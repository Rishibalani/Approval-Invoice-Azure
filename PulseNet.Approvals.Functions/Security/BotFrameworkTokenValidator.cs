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
/// well-known OpenID metadata endpoint (TeamsBot:OpenIdMetadataUrl), which the
/// configuration manager below fetches and caches, refreshing on its own
/// schedule. Pinning a key would work until the day it rotated and then fail
/// silently at 3am.
///
/// Accepted issuers (TeamsBot:ValidTokenIssuers) and clock skew
/// (TeamsBot:TokenClockSkewSeconds) are configuration too; the settings are
/// validated at startup whenever Teams runs in Bot mode.
/// </summary>
public sealed class BotFrameworkTokenValidator
{
    private const string BearerPrefix = "Bearer ";

    /// <summary>
    /// Null when TeamsBot:OpenIdMetadataUrl is not configured - the bot is not
    /// in use, and every call fails closed rather than throwing at construction.
    /// </summary>
    private readonly IConfigurationManager<OpenIdConnectConfiguration>? _configManager;
    private readonly TeamsBotOptions _options;
    private readonly ILogger<BotFrameworkTokenValidator> _logger;
    private readonly JwtSecurityTokenHandler _handler = new();

    public BotFrameworkTokenValidator(
        IOptions<TeamsBotOptions> options,
        ILogger<BotFrameworkTokenValidator> logger)
    {
        _options = options.Value;
        _logger = logger;

        if (!string.IsNullOrWhiteSpace(_options.OpenIdMetadataUrl))
        {
            _configManager = new ConfigurationManager<OpenIdConnectConfiguration>(
                _options.OpenIdMetadataUrl,
                new OpenIdConnectConfigurationRetriever(),
                new HttpDocumentRetriever { RequireHttps = true });
        }
    }

    public async Task<bool> IsValidAsync(string? authorizationHeader, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(_options.AppId) || _configManager is null)
        {
            // Fail closed. Missing configuration must never mean "skip the check".
            _logger.LogError(
                "TeamsBot:AppId or TeamsBot:OpenIdMetadataUrl is not configured. Refusing all bot traffic.");
            return false;
        }

        if (string.IsNullOrWhiteSpace(authorizationHeader) ||
            !authorizationHeader.StartsWith(BearerPrefix, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var token = authorizationHeader[BearerPrefix.Length..].Trim();

        try
        {
            var config = await _configManager.GetConfigurationAsync(cancellationToken);

            var parameters = new TokenValidationParameters
            {
                ValidateIssuer = true,

                // The Bot Framework issuer plus, for single-tenant bots, their
                // own tenant's issuers - {tenantId} substituted from TenantId.
                ValidIssuers = _options.ResolvedTokenIssuers,

                // The check that actually matters: a valid token issued for a
                // different bot must not be accepted here.
                ValidateAudience = true,
                ValidAudience = _options.AppId,

                ValidateIssuerSigningKey = true,
                IssuerSigningKeys = config.SigningKeys,

                ValidateLifetime = true,
                ClockSkew = TimeSpan.FromSeconds(_options.TokenClockSkewSeconds)
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
