using System.ComponentModel.DataAnnotations;

namespace PulseNet.Approvals.Functions.Options;

/// <summary>
/// Azure Bot Service credentials and Teams delivery settings.
///
/// These come from Parts 1 and 2 of the runbook: the Entra app registration
/// that represents the bot, and the Azure Bot resource that connects it to the
/// Teams channel.
///
/// VALIDATION
///
/// The HttpClient timeouts carry DataAnnotations and are always required: the
/// typed clients are constructed whenever the Teams sender is resolved, bot or
/// not. Everything else is required only when Channels:Teams:DeliveryMode is
/// Bot - see Options/Validation/TeamsBotOptionsValidator.
/// </summary>
public sealed class TeamsBotOptions
{
    public const string SectionName = "TeamsBot";

    /// <summary>
    /// Placeholder accepted in ValidTokenIssuers, replaced with TenantId at the
    /// point of use so the tenant GUID is configured once.
    /// </summary>
    public const string TenantIdPlaceholder = "{tenantId}";

    /// <summary>
    /// Microsoft App ID from the Azure Bot resource. Same GUID as the Entra
    /// app registration's Application (client) ID.
    /// </summary>
    public string AppId { get; set; } = string.Empty;

    /// <summary>Client secret from that registration.</summary>
    public string AppPassword { get; set; } = string.Empty;

    /// <summary>
    /// Directory (tenant) ID. Required because bots are now registered
    /// SingleTenant - Microsoft retired MultiTenant for new bots on
    /// 31 July 2025, and the token endpoint differs between the two.
    /// </summary>
    public string TenantId { get; set; } = string.Empty;

    /// <summary>
    /// Teams service URL for your region. India is /in/, Americas /amer/,
    /// Europe /emea/.
    ///
    /// Used only when no stored conversation reference exists. The real
    /// service URL arrives on every inbound activity and is stored with the
    /// conversation reference, because it can change and the stored one is
    /// always more trustworthy than configuration.
    /// </summary>
    public string DefaultServiceUrl { get; set; } = string.Empty;

    /// <summary>
    /// The "id" from your Teams manifest. Optional: only the provisioning
    /// endpoint's install step needs it, and it reports clearly when unset.
    ///
    /// NOT the catalogue app ID - Teams assigns a separate one when the package
    /// is uploaded, and the install call wants that. GraphDirectoryClient looks
    /// it up from this value. Passing the wrong one produces a 404 that reads
    /// as though the app does not exist.
    /// </summary>
    public string ManifestId { get; set; } = string.Empty;

    /// <summary>
    /// Table holding which Teams message carried which approval, so a card
    /// can be replaced once a decision is made anywhere.
    /// </summary>
    public string SentCardTable { get; set; } = string.Empty;

    /// <summary>Table Storage table holding conversation references.</summary>
    public string ConversationTable { get; set; } = string.Empty;

    /// <summary>
    /// When true, the sender attempts to create a 1:1 conversation from the
    /// approver's Entra object ID even with no stored reference.
    ///
    /// This only succeeds if the Teams app is already installed for that user.
    /// Teams refuses to create a conversation with a bot the user has never
    /// added - there is no way around that short of proactive installation,
    /// which needs TeamsAppInstallation.ReadWriteForUser.All.
    ///
    /// Must be present in configuration when the bot is in use.
    /// </summary>
    public bool AttemptDirectConversation { get; set; }

    /// <summary>
    /// Microsoft Graph credentials for resolving a UPN to an Entra object ID.
    /// Usually the same registration as the bot, with User.Read.All added.
    /// Leave blank to fall back to the bot's own credentials (AppId /
    /// AppPassword), or to skip per-user Graph lookups and rely on cached
    /// object IDs.
    /// </summary>
    public string GraphClientId { get; set; } = string.Empty;
    public string GraphClientSecret { get; set; } = string.Empty;

    // ------------------------------------------------------------------
    //  Endpoints and scopes
    // ------------------------------------------------------------------

    /// <summary>
    /// OAuth scope for Bot Connector calls, e.g. https://api.botframework.com/.default
    /// </summary>
    public string BotFrameworkScope { get; set; } = string.Empty;

    /// <summary>
    /// OpenID metadata document publishing the Bot Framework signing keys, e.g.
    /// https://login.botframework.com/v1/.well-known/openidconfiguration
    /// </summary>
    public string OpenIdMetadataUrl { get; set; } = string.Empty;

    /// <summary>
    /// Comma-separated accepted iss claims on inbound Bot Framework tokens.
    /// {tenantId} is replaced with TenantId, because single-tenant bots see
    /// their own tenant as the issuer.
    /// </summary>
    public string ValidTokenIssuers { get; set; } = string.Empty;

    /// <summary>Clock skew allowed when validating an inbound Bot Framework token.</summary>
    public int TokenClockSkewSeconds { get; set; }

    /// <summary>
    /// Microsoft Graph root including the API version, e.g.
    /// https://graph.microsoft.com/v1.0
    /// </summary>
    public string GraphBaseUrl { get; set; } = string.Empty;

    /// <summary>OAuth scope for Graph calls, e.g. https://graph.microsoft.com/.default</summary>
    public string GraphScope { get; set; } = string.Empty;

    /// <summary>
    /// Users per page when the provisioning endpoint reads the directory.
    /// Graph caps $top at 999 for /users.
    /// </summary>
    public int GraphDirectoryPageSize { get; set; }

    /// <summary>
    /// Pause after Graph throttles an app install during provisioning.
    /// Backing off beats hammering it and having the rest of the run fail too.
    /// </summary>
    public int ProvisioningThrottleDelaySeconds { get; set; }

    // ------------------------------------------------------------------
    //  HttpClient timeouts - always required
    // ------------------------------------------------------------------

    /// <summary>HttpClient timeout for Bot Connector (and per-user Graph) calls.</summary>
    [Range(1, int.MaxValue, ErrorMessage = "TeamsBot__HttpTimeoutSeconds must be > 0.")]
    public int HttpTimeoutSeconds { get; set; }

    /// <summary>
    /// HttpClient timeout for bulk directory reads and app installs. Generous:
    /// a directory page of 999 users on a slow tenant.
    /// </summary>
    [Range(1, int.MaxValue, ErrorMessage = "TeamsBot__GraphHttpTimeoutSeconds must be > 0.")]
    public int GraphHttpTimeoutSeconds { get; set; }

    /// <summary>ValidTokenIssuers split, with {tenantId} substituted.</summary>
    public string[] ResolvedTokenIssuers =>
        ValidTokenIssuers
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(i => i.Replace(TenantIdPlaceholder, TenantId, StringComparison.OrdinalIgnoreCase))
            .ToArray();
}
