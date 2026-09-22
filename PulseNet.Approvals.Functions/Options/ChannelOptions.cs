namespace PulseNet.Approvals.Functions.Options;

using PulseNet.Approvals.Functions.Models;

/// <summary>
/// Per-channel configuration. Mirrors the fields on the Business Central setup
/// page so the two stay in step, and so a channel can be switched off from
/// either end.
///
/// Business Central is authoritative for POLICY (may this be approved in a
/// channel at all). This class is authoritative for TRANSPORT (how do we
/// physically reach the channel). Keeping that line clean is what stops
/// financial rules leaking out of the system of record.
///
/// VALIDATION
///
/// Nothing here has a code-level default. DataAnnotations do not recurse into
/// the nested channel classes, so every rule - including the conditional ones
/// (WhatsApp settings only when WhatsApp is enabled, the webhook URL only in
/// WorkflowWebhook mode, and so on) - lives in
/// Options/Validation/ChannelOptionsValidator and runs at startup.
/// </summary>
public sealed class ChannelOptions
{
    public const string SectionName = "Channels";

    public TeamsChannelOptions Teams { get; set; } = new();
    public OutlookChannelOptions Outlook { get; set; } = new();
    public WhatsAppChannelOptions WhatsApp { get; set; } = new();

    /// <summary>
    /// FULL URL of the action endpoint that Link-mode buttons point at,
    /// including the route, e.g.
    /// https://pulsenet-approvals.azurewebsites.net/api/approvals/act
    ///
    /// Used verbatim - only the ?t= token parameter is appended - so it must
    /// match the ApprovalAct function's route. Required when any enabled
    /// channel uses Link action mode.
    /// </summary>
    public string ActionEndpointBaseUrl { get; set; } = string.Empty;

    /// <summary>
    /// Channel tried when every enabled channel fails and Business Central's
    /// payload does not name a fallback of its own. Outlook is the usual
    /// choice: every approver has a mailbox, which is not true of a Teams app
    /// install. Must be present in configuration.
    /// </summary>
    public ApprovalChannel FallbackChannel { get; set; }

    /// <summary>
    /// Currency code shown when Business Central sends neither a document
    /// Currency Code nor an LCY code - normally the company's home currency.
    /// A bare "10.00" tells the approver nothing.
    /// </summary>
    public string FallbackCurrencyCode { get; set; } = string.Empty;

    /// <summary>
    /// Document lines rendered on a card before "+N more line(s)". Teams
    /// rejects cards over 28 KB, so keep this modest (10 was the old constant).
    /// Business Central's "Email Max Lines" should normally match it.
    /// </summary>
    public int MaxLinesOnCard { get; set; }
}

public sealed class TeamsChannelOptions
{
    /// <summary>Must be present in configuration. Disabled switches Teams off.</summary>
    public ChannelDeliveryMode DeliveryMode { get; set; }

    /// <summary>Must be present in configuration.</summary>
    public ChannelActionMode ActionMode { get; set; }

    /// <summary>
    /// Workflows incoming webhook URL. Posts to a channel or a chat depending
    /// on which Workflows template created it.
    ///
    /// A CHANNEL webhook is visible to everyone in that channel, so anyone
    /// there could tap Approve. That is acceptable in a sandbox and is why
    /// ActionEndpoint sign-in enforcement exists - see ActionTokenOptions.
    /// </summary>
    public string WebhookUrl { get; set; } = string.Empty;

    /// <summary>
    /// What shape the webhook expects.
    ///
    /// FlowRouted - a Power Automate flow that takes the recipient as a
    ///   parameter and posts a 1:1 chat via Flow bot. Per-approver delivery,
    ///   no bot registration. The usual choice.
    ///
    /// TeamsMessage - the raw Teams message envelope, posted straight into a
    ///   fixed channel. One destination for everyone, so it cannot address an
    ///   individual approver. Useful for a shared ops channel or a quick test.
    ///
    /// Must be present in configuration when DeliveryMode is WorkflowWebhook.
    /// </summary>
    public WebhookPayloadMode PayloadMode { get; set; }

    /// <summary>HttpClient timeout for posts to the Workflows webhook. Always required.</summary>
    public int WebhookHttpTimeoutSeconds { get; set; }

    /// <summary>
    /// Full-width cards. A desktop nicety and a known source of mobile
    /// rendering trouble, so normally false. Turn on only if the phones in use
    /// are known to handle it. Must be present when Teams is enabled.
    /// </summary>
    public bool UseFullWidthCard { get; set; }

    /// <summary>
    /// The Adaptive Card `refresh` block, which lets a card re-fetch itself
    /// when reopened.
    ///
    /// Normally false. It is Adaptive Card 1.4, and a mobile client that
    /// cannot parse it may drop the WHOLE CARD rather than just the refresh -
    /// which renders as an empty message with a timestamp and no error
    /// anywhere. Must be present when Teams is enabled.
    ///
    /// Cards now update themselves through CardRefreshService when a decision
    /// lands, so this is a secondary path rather than the only one.
    /// </summary>
    public bool UseCardRefreshBlock { get; set; }

    // Bot mode. Unused until DeliveryMode is Bot.
    public string BotAppId { get; set; } = string.Empty;
    public string BotAppPassword { get; set; } = string.Empty;
    public string BotServiceUrl { get; set; } = string.Empty;
}

public enum WebhookPayloadMode
{
    /// <summary>{ recipientUpn, summary, cardJson } - flow routes to a 1:1 chat.</summary>
    FlowRouted = 0,

    /// <summary>{ type: "message", attachments: [...] } - posts to a fixed channel.</summary>
    TeamsMessage = 1
}

/// <summary>
/// Outlook transport settings. The Outlook senders are preserved but not
/// registered; the settings below are required only when DeliveryMode is not
/// Disabled. RequireRejectionReason is the exception - it is read by the Teams
/// bot path too, so it is always required.
/// </summary>
public sealed class OutlookChannelOptions
{
    /// <summary>Must be present in configuration. Disabled switches Outlook off.</summary>
    public ChannelDeliveryMode DeliveryMode { get; set; }

    /// <summary>Must be present in configuration.</summary>
    public ChannelActionMode ActionMode { get; set; }

    public string FromAddress { get; set; } = string.Empty;
    public string FromDisplayName { get; set; } = string.Empty;

    /// <summary>
    /// Which transport actually puts the mail on the wire (None, Graph, Acs).
    /// A real decision with cost and permission consequences - see
    /// PlainEmailSender for the three options.
    /// </summary>
    public string Transport { get; set; } = string.Empty;

    /// <summary>Azure Communication Services connection string, when Transport is Acs.</summary>
    public string AcsConnectionString { get; set; } = string.Empty;

    // ------------------------------------------------------------------
    //  Graph send-mail transport
    //  Used by GraphEmailTransport. Application permission Mail.Send, granted
    //  with admin consent, and scoped with an ApplicationAccessPolicy so the
    //  app can only send as FromAddress - without that policy Mail.Send lets
    //  it send as ANY mailbox in the tenant.
    // ------------------------------------------------------------------

    public string TenantId { get; set; } = string.Empty;
    public string ClientId { get; set; } = string.Empty;

    /// <summary>
    /// Keep in Key Vault in every environment above local. A client secret in
    /// app settings is readable by anyone with Reader on the Function App.
    /// </summary>
    public string ClientSecret { get; set; } = string.Empty;

    // ------------------------------------------------------------------
    //  Actionable Messages inbound validation
    //  Used by ActionableMessageTokenValidator / OutlookActionFunction.
    //
    //  When a user taps Approve in an Outlook actionable message, Microsoft
    //  POSTs to the action endpoint with a bearer token it minted. That token
    //  is the ONLY proof the click came from Outlook and not from anyone who
    //  guessed the URL, so validation is not optional in production.
    // ------------------------------------------------------------------

    /// <summary>Actionable Messages originator ID from the provider registration.</summary>
    public string OriginatorId { get; set; } = string.Empty;

    /// <summary>
    /// False only for local development, where no real Outlook token exists.
    /// Fails the build of a production deployment if left false - assert it in
    /// Program.cs rather than trusting configuration review. Must be present
    /// when Outlook is enabled.
    /// </summary>
    public bool ValidateInboundToken { get; set; }

    /// <summary>
    /// OpenID metadata document for the Actionable Messages signing keys, e.g.
    /// https://substrate.office.com/sts/common/.well-known/openid-configuration
    /// </summary>
    public string TokenMetadataUrl { get; set; } = string.Empty;

    /// <summary>
    /// Comma-separated accepted iss claims, e.g. https://substrate.office.com/sts/
    /// Split at the point of use.
    /// </summary>
    public string ValidTokenIssuers { get; set; } = string.Empty;

    /// <summary>
    /// The aud claim, which Microsoft sets to the ORIGIN of your action
    /// endpoint - scheme and host, no path. If this does not match exactly,
    /// every action returns 401 and the cause is not obvious from the logs.
    /// e.g. https://pulsenet-approvals.azurewebsites.net
    /// </summary>
    public string ExpectedTokenAudience { get; set; } = string.Empty;

    /// <summary>
    /// Require the token's sub claim to equal the approver the card was sent
    /// to. Stops a forwarded email being actioned by the recipient: the mail
    /// forwards, the token does not re-mint for the new reader, but a shared
    /// mailbox can still produce a surprise. Leave true. Must be present when
    /// Outlook is enabled.
    /// </summary>
    public bool RequireMailboxMatch { get; set; }

    /// <summary>
    /// Refuse a rejection that carries no reason.
    ///
    /// Enforced server-side because there is nowhere else it CAN be enforced.
    /// Adaptive Card 1.0 accepts isRequired and ignores it, so the card cannot
    /// stop an empty submission - and a client-side check would not stop a
    /// crafted request in any case.
    ///
    /// Read by the Teams bot path as well as Outlook. The rule is one policy,
    /// not one per channel, so a single switch beats several that can disagree
    /// about whether a reason is needed. Always required, whatever Outlook's
    /// delivery mode - a missing key must never silently mean "no reason".
    /// </summary>
    public bool RequireRejectionReason { get; set; }
}

/// <summary>
/// WhatsApp Cloud API settings. Everything except DeliveryMode and the
/// HttpClient timeout is required only when DeliveryMode is not Disabled.
/// </summary>
public sealed class WhatsAppChannelOptions
{
    /// <summary>Must be present in configuration. Disabled switches WhatsApp off.</summary>
    public ChannelDeliveryMode DeliveryMode { get; set; }

    public ChannelActionMode ActionMode { get; set; }

    public string PhoneNumberId { get; set; } = string.Empty;
    public string AccessToken { get; set; } = string.Empty;

    /// <summary>Must match the template exactly as approved by Meta.</summary>
    public string TemplateName { get; set; } = string.Empty;

    /// <summary>Language code the template was approved under, e.g. en.</summary>
    public string TemplateLanguage { get; set; } = string.Empty;
    /// <summary>
    /// The Meta app secret. Used to verify X-Hub-Signature-256 on every
    /// inbound webhook, which is the only thing standing between a public URL
    /// and a forged approval.
    /// </summary>
    public string AppSecret { get; set; } = string.Empty;

    /// <summary>
    /// A string you invent. Meta echoes it during webhook verification and the
    /// endpoint refuses anything else.
    /// </summary>
    public string WebhookVerifyToken { get; set; } = string.Empty;

    /// <summary>
    /// Meta Graph API root, without the version, e.g. https://graph.facebook.com
    /// </summary>
    public string GraphApiBaseUrl { get; set; } = string.Empty;

    /// <summary>
    /// Graph API version in the send URL, e.g. v21.0. Pinned rather than
    /// floating - Meta deprecates versions on a schedule and a silent bump is
    /// not something to discover from a production failure.
    /// </summary>
    public string ApiVersion { get; set; } = string.Empty;

    /// <summary>HttpClient timeout for Cloud API calls. Always required.</summary>
    public int HttpTimeoutSeconds { get; set; }

    /// <summary>
    /// Table holding rejections waiting for their reason. See
    /// PendingRejectionStore for why WhatsApp needs this and the others do not.
    /// </summary>
    public string PendingRejectionTable { get; set; } = string.Empty;

    /// <summary>
    /// How long to wait for a rejection reason before giving up.
    ///
    /// Shorter than the action token TTL on purpose: the approver has already
    /// tapped Reject, so they are present and typing. It should be generous for
    /// someone mid-conversation and short enough that a forgotten tap does not
    /// leave a rejection armed for long.
    /// </summary>
    public int RejectionReasonTimeoutMinutes { get; set; }

    /// <summary>
    /// Inbound message IDs, for deduplication. Meta retries a webhook for up
    /// to 24 hours if it does not get a 200, so without this a slow response
    /// can approve the same invoice twice.
    /// </summary>
    public string InboundDedupeTable { get; set; } = string.Empty;
}
