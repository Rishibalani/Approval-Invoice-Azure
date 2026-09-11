using System.IdentityModel.Tokens.Jwt;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Protocols;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Tokens;
using PulseNet.Approvals.Functions.Options;

namespace PulseNet.Approvals.Functions.Security;

/// <summary>
/// Verifies that an Action.Http POST genuinely came from Outlook.
///
/// WHY THIS MATTERS
///
/// The endpoint is a public URL with no function key - Outlook cannot attach
/// one. Without this check, anyone who found the URL could post a forged
/// approval. The action token limits the damage (it is signed, single-use and
/// short-lived) but the transport should not be open either.
///
/// THE AUTH MODEL CHANGED IN JUNE 2026
///
/// Legacy External Access Tokens for Actionable Messages were retired on
/// 8 June 2026. The replacement is a Microsoft Entra ID issued token,
/// validated against Entra's published signing keys rather than the old
/// substrate endpoint.
///
/// Every sample and blog post describing a "Microsoft-issued bearer token"
/// with a validation URL under substrate.office.com predates that change.
///
/// This is recent enough that the exact issuer and audience values are worth
/// confirming against current Microsoft documentation for your tenant before
/// production - both are configurable below rather than hard-coded, precisely
/// so that confirming them is an app-setting change and not a redeploy.
///
/// WHAT THE TOKEN GIVES US BEYOND AUTHENTICITY
///
/// The sender and recipient claims. That is a genuine identity signal - it
/// asserts the MAILBOX the action came from. Weaker than an interactive
/// sign-in, because it does not prove a person was present, but strong enough
/// to check against the approver the action token was minted for. That check
/// is what stops a forwarded email being actioned by the wrong person.
/// </summary>
public sealed class ActionableMessageTokenValidator
{
    private readonly IConfigurationManager<OpenIdConnectConfiguration> _configManager;
    private readonly OutlookChannelOptions _options;
    private readonly ILogger<ActionableMessageTokenValidator> _logger;
    private readonly JwtSecurityTokenHandler _handler = new();

    public ActionableMessageTokenValidator(
        IOptions<ChannelOptions> options,
        ILogger<ActionableMessageTokenValidator> logger)
    {
        _options = options.Value.Outlook;
        _logger = logger;

        _configManager = new ConfigurationManager<OpenIdConnectConfiguration>(
            _options.TokenMetadataUrl,
            new OpenIdConnectConfigurationRetriever(),
            new HttpDocumentRetriever { RequireHttps = true });
    }

    public async Task<ActionableMessageValidation> ValidateAsync(
        string? authorizationHeader,
        CancellationToken cancellationToken)
    {
        // Fail closed. Turning validation off is a development convenience and
        // must never be the default - an unauthenticated approval endpoint is
        // exactly the thing this class exists to prevent.
        if (!_options.ValidateInboundToken)
        {
            _logger.LogWarning(
                "Outlook token validation is DISABLED. Acceptable locally, never in a reachable environment.");

            return ActionableMessageValidation.Valid(null);
        }

        if (string.IsNullOrWhiteSpace(authorizationHeader) ||
            !authorizationHeader.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
        {
            return ActionableMessageValidation.Invalid("missing_bearer_token");
        }

        var token = authorizationHeader["Bearer ".Length..].Trim();

        try
        {
            var config = await _configManager.GetConfigurationAsync(cancellationToken);

            var parameters = new TokenValidationParameters
            {
                ValidateIssuer = true,
                ValidIssuers = _options.ValidTokenIssuers
                    .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries),

                // The audience is the action endpoint URL. A token minted for
                // somebody else's endpoint must not be accepted here, and
                // checking only the issuer is a common and serious mistake.
                ValidateAudience = true,
                ValidAudience = _options.ExpectedTokenAudience,

                ValidateIssuerSigningKey = true,
                IssuerSigningKeys = config.SigningKeys,

                ValidateLifetime = true,
                ClockSkew = TimeSpan.FromMinutes(5)
            };

            _handler.ValidateToken(token, parameters, out var validated);

            var jwt = validated as JwtSecurityToken;

            // Claim names vary between token versions, so several are checked.
            // Returning null rather than throwing when none is present keeps
            // the mailbox check optional - it is a defence in depth, not the
            // primary control.
            var sender = jwt?.Claims.FirstOrDefault(claim =>
                claim.Type is "sender" or "sub" or "upn" or "preferred_username")?.Value;

            return ActionableMessageValidation.Valid(sender);
        }
        catch (SecurityTokenException ex)
        {
            // Expected when somebody probes the endpoint. Information, not error.
            _logger.LogInformation("Outlook token rejected: {Reason}", ex.Message);
            return ActionableMessageValidation.Invalid("token_invalid");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Outlook token validation failed unexpectedly.");
            return ActionableMessageValidation.Invalid("validation_error");
        }
    }
}

public sealed record ActionableMessageValidation
{
    public required bool IsValid { get; init; }
    public string? FailureReason { get; init; }

    /// <summary>
    /// The mailbox the action came from, when the token carries it. Compared
    /// against the approver the action token was minted for.
    /// </summary>
    public string? SenderEmail { get; init; }

    public static ActionableMessageValidation Invalid(string reason) =>
        new() { IsValid = false, FailureReason = reason };

    public static ActionableMessageValidation Valid(string? senderEmail) =>
        new() { IsValid = true, SenderEmail = senderEmail };
}
