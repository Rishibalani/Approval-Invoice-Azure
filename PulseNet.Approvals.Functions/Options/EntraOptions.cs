using System.ComponentModel.DataAnnotations;

namespace PulseNet.Approvals.Functions.Options;

/// <summary>
/// Microsoft Entra ID (identity platform) settings shared by every client that
/// acquires a client-credentials token: Business Central, the Bot Connector
/// and Microsoft Graph.
///
/// One section rather than one copy per client, so a sovereign-cloud move or a
/// change to the renewal margin is one setting, not three that can disagree.
/// </summary>
public sealed class EntraOptions
{
    public const string SectionName = "Entra";

    /// <summary>
    /// The v2.0 token endpoint path under {AuthorityHost}/{tenantId}/. Part of
    /// the OAuth protocol surface of the identity platform, not a deployment
    /// choice - the host is what varies between clouds.
    /// </summary>
    private const string TokenPath = "oauth2/v2.0/token";

    /// <summary>
    /// Entra authority host, e.g. https://login.microsoftonline.com
    /// </summary>
    [Required(ErrorMessage = "Entra__AuthorityHost is required.")]
    [Url(ErrorMessage = "Entra__AuthorityHost must be an absolute URL.")]
    public string AuthorityHost { get; set; } = string.Empty;

    /// <summary>
    /// How long before expiry a cached access token is renewed. A token
    /// expiring between our check and the resource server's validation
    /// produces a 401 indistinguishable from a misconfiguration; renewing
    /// early makes that race impossible.
    /// </summary>
    [Range(1, int.MaxValue, ErrorMessage = "Entra__TokenRefreshSkewSeconds must be > 0.")]
    public int TokenRefreshSkewSeconds { get; set; }

    public TimeSpan TokenRefreshSkew => TimeSpan.FromSeconds(TokenRefreshSkewSeconds);

    /// <summary>Client-credentials token endpoint for a tenant.</summary>
    public string TokenEndpoint(string tenantId) =>
        $"{AuthorityHost.TrimEnd('/')}/{tenantId}/{TokenPath}";
}
