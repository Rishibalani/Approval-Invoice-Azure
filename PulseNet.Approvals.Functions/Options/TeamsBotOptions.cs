namespace PulseNet.Approvals.Functions.Options;

/// <summary>
/// Azure Bot Service credentials and Teams delivery settings.
///
/// These come from Parts 1 and 2 of the runbook: the Entra app registration
/// that represents the bot, and the Azure Bot resource that connects it to the
/// Teams channel.
/// </summary>
public sealed class TeamsBotOptions
{
    public const string SectionName = "TeamsBot";

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
    /// This is a fallback only. The real service URL arrives on every inbound
    /// activity and is stored with the conversation reference, because it can
    /// change and the stored one is always more trustworthy than a guess.
    /// </summary>
    public string DefaultServiceUrl { get; set; } = "https://smba.trafficmanager.net/in/";

    /// <summary>
    /// The "id" from your Teams manifest.
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
    public string SentCardTable { get; set; } = "teamssentcards";

    /// <summary>Table Storage table holding conversation references.</summary>
    public string ConversationTable { get; set; } = "teamsconversations";

    /// <summary>
    /// When true, the sender attempts to create a 1:1 conversation from the
    /// approver's Entra object ID even with no stored reference.
    ///
    /// This only succeeds if the Teams app is already installed for that user.
    /// Teams refuses to create a conversation with a bot the user has never
    /// added - there is no way around that short of proactive installation,
    /// which needs TeamsAppInstallation.ReadWriteForUser.All.
    /// </summary>
    public bool AttemptDirectConversation { get; set; } = true;

    /// <summary>
    /// Microsoft Graph credentials for resolving a UPN to an Entra object ID.
    /// Usually the same registration as the bot, with User.Read.All added.
    /// Leave blank to skip Graph entirely and rely on cached object IDs.
    /// </summary>
    public string GraphClientId { get; set; } = string.Empty;
    public string GraphClientSecret { get; set; } = string.Empty;

    public const string BotFrameworkScope = "https://api.botframework.com/.default";
    public const string GraphScope = "https://graph.microsoft.com/.default";
}
