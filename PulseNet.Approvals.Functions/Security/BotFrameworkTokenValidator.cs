using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Protocols;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Tokens;
using PulseNet.Approvals.Functions.Options;
using System.IdentityModel.Tokens.Jwt;

namespace PulseNet.Approvals.Functions.Security;

/// <summary>
/// Verifies that a call to the bot endpoint really came from the Bot Framework.
///
/// WHY THIS MATTERS MORE THAN IT LOOKS
///
/// The messaging endpoint is a public URL with no function key - Teams cannot
/// attach one. Without this check, anyone who found the URL could post a
/// forged invoke and approve invoices.
///
/// The JWT is issued by Microsoft, expires, and names the bot it was issued
/// for. Validating the audience against our own App ID is the part that
/// matters: a valid token for somebody else's bot must not be accepted here.
///
/// SIGNING KEYS ARE FETCHED, NOT CONFIGURED
///
/// Microsoft rotates the Bot Framework signing keys. They are published at a
/// well-known OpenID metadata endpoint, which the configuration manager below
/// fetches and caches, refreshing on its own schedule. Pinning a key would
/// work until the day it rotated and then fail silently at 3am.
/// </summary>
public sealed class BotFrameworkTokenValidator
{
    private const string MetadataUrl =
        "https://login.botframework.com/v1/.well-known/openidconfiguration";

    private const string ExpectedIssuer = "https://api.botframework.com";

    private readonly IConfigurationManager<OpenIdConnectConfiguration> _configManager;
    private readonly TeamsBotOptions _options;
    private readonly ILogger<BotFrameworkTokenValidator> _logger;
    private readonly JwtSecurityTokenHandler _handler = new();

    public BotFrameworkTokenValidator(
        IOptions<TeamsBotOptions> options,
        ILogger<BotFrameworkTokenValidator> logger)
    {
        _options = options.Value;
        _logger = logger;

        _configManager = new ConfigurationManager<OpenIdConnectConfiguration>(
            MetadataUrl,
            new OpenIdConnectConfigurationRetriever(),
            new HttpDocumentRetriever { RequireHttps = true });
    }

    public async Task<bool> IsValidAsync(string? authorizationHeader, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(_options.AppId))
        {
            // Fail closed. A missing App ID must never mean "skip the check".
            _logger.LogError("TeamsBot:AppId is not configured. Refusing all bot traffic.");
            return false;
        }

        if (string.IsNullOrWhiteSpace(authorizationHeader) ||
            !authorizationHeader.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var token = authorizationHeader["Bearer ".Length..].Trim();

        try
        {
            var config = await _configManager.GetConfigurationAsync(cancellationToken);

            var parameters = new TokenValidationParameters
            {
                ValidateIssuer = true,
                ValidIssuers = new[]
                {
                    ExpectedIssuer,
                    // Single-tenant bots see their own tenant as the issuer.
                    $"https://login.microsoftonline.com/{_options.TenantId}/v2.0",
                    $"https://sts.windows.net/{_options.TenantId}/"
                },

                // The check that actually matters: a valid token issued for a
                // different bot must not be accepted here.
                ValidateAudience = true,
                ValidAudience = _options.AppId,

                ValidateIssuerSigningKey = true,
                IssuerSigningKeys = config.SigningKeys,

                ValidateLifetime = true,
                ClockSkew = TimeSpan.FromMinutes(5)
            };

            _handler.ValidateToken(token, parameters, out _);
            return true;
        }
        catch (SecurityTokenException ex)
        {
            // Expected when someone probes the endpoint. Information, not error.
            _logger.LogInformation("Bot token rejected: {Reason}", ex.Message);
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Bot token validation failed unexpectedly.");
            return false;
        }
    }
}
