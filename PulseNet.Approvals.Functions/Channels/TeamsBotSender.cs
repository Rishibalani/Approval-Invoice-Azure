using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Cards;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Options;
using PulseNet.Approvals.Functions.Services;

namespace PulseNet.Approvals.Functions.Channels;

/// <summary>
/// Delivers an approval card to the approver's own 1:1 Teams chat.
///
/// THE DELIVERY LADDER
///
///   1. Stored conversation reference  - captured when they installed the app
///   2. Create a conversation          - works if installed but never captured
///   3. Return failure                 - dispatcher falls back to Outlook
///
/// Step 2 exists for the people who added the app before this code shipped.
/// It cannot rescue someone who has never installed it: Teams refuses to open
/// a chat with a bot the user has not added, and no permission short of
/// TeamsAppInstallation.ReadWriteForUser.All changes that.
///
/// So an approver who has not installed the Teams app will always fall through
/// to email. That is a rollout task, not a bug, and it is why the fallback
/// channel matters more than it looks.
///
/// WHY THIS BEATS THE WEBHOOK
///
/// Identity. When a card is posted to a channel, everyone in it can tap
/// Approve and the token is all that binds the action to a person. A 1:1 chat
/// reaches only the approver, and Action.Execute makes Teams assert their
/// verified Entra identity on the tap itself - no browser, no sign-in, no
/// shared link.
///
/// Cards also update in place, so a decided approval stops looking live.
/// </summary>
public sealed class TeamsBotSender : IChannelSender
{
    private readonly BotConnectorClient _connector;
    private readonly ConversationReferenceStore _conversations;
    private readonly ApprovalCardBuilder _cardBuilder;
    private readonly TeamsBotOptions _options;
    private readonly ILogger<TeamsBotSender> _logger;

    public ChannelDeliveryMode Mode => ChannelDeliveryMode.Bot;
    public ApprovalChannel Channel => ApprovalChannel.Teams;

    public TeamsBotSender(
        BotConnectorClient connector,
        ConversationReferenceStore conversations,
        ApprovalCardBuilder cardBuilder,
        IOptions<TeamsBotOptions> options,
        ILogger<TeamsBotSender> logger)
    {
        _connector = connector;
        _conversations = conversations;
        _cardBuilder = cardBuilder;
        _options = options.Value;
        _logger = logger;
    }

    public async Task<ChannelSendResult> SendAsync(
        ApprovalDispatchPayload payload,
        ChannelActionMode actionMode,
        string? approveUrl,
        string? rejectUrl,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(_options.AppId))
        {
            return ChannelSendResult.Fail(Channel, "bot_not_configured");
        }

        // ---- Who are we sending to ------------------------------------
        var aadObjectId = await ResolveObjectIdAsync(payload, cancellationToken);

        if (aadObjectId is null)
        {
            return ChannelSendResult.Fail(Channel, "could_not_resolve_approver");
        }

        // ---- How do we reach them -------------------------------------
        var (conversationId, serviceUrl) = await ResolveConversationAsync(aadObjectId, cancellationToken);

        if (conversationId is null)
        {
            _logger.LogInformation(
                "No Teams conversation available for {Approver}. Falling back to another channel.",
                payload.Approver.UserId);

            return ChannelSendResult.Fail(Channel, "teams_app_not_installed");
        }

        // ---- Build and send -------------------------------------------
        try
        {
            var card = _cardBuilder.Build(payload, actionMode, approveUrl, rejectUrl);
            var summary = BuildSummary(payload);

            var activityId = await _connector.SendCardAsync(
                conversationId, serviceUrl, card, summary, cancellationToken);

            if (activityId is null)
            {
                return ChannelSendResult.Fail(Channel, "send_failed", transient: true);
            }

            _logger.LogInformation(
                "Teams card sent to {Approver} for {DocumentNo}, activity {ActivityId}.",
                payload.Approver.UserId, payload.Document.DocumentNo, activityId);

            // The activity ID is what makes an in-place update possible when
            // the approval is decided. Composed with the conversation ID
            // because updating needs both.
            return ChannelSendResult.Ok(Channel, $"{conversationId}|{activityId}");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Teams bot send threw for {DocumentNo}.", payload.Document.DocumentNo);
            return ChannelSendResult.Fail(Channel, "bot_send_exception", transient: true);
        }
    }

    // ------------------------------------------------------------------

    /// <summary>
    /// Entra object ID for the approver, cheapest source first.
    ///
    /// Business Central caches it on the identity record after the first
    /// lookup, so Graph is normally consulted once per approver ever. That
    /// matters because Graph is the only part of this path needing a
    /// permission somebody has to consent to.
    /// </summary>
    private async Task<string?> ResolveObjectIdAsync(
        ApprovalDispatchPayload payload,
        CancellationToken cancellationToken)
    {
        // 1. Cached in Business Central and sent with the payload.
        if (!string.IsNullOrWhiteSpace(payload.Approver.EntraObjectId))
        {
            return payload.Approver.EntraObjectId;
        }

        // 2. Resolve from the UPN via Graph.
        if (!string.IsNullOrWhiteSpace(payload.Approver.Upn))
        {
            var resolved = await _connector.TryResolveObjectIdAsync(payload.Approver.Upn, cancellationToken);

            if (resolved is not null)
            {
                _logger.LogInformation(
                    "Resolved {Upn} to an Entra object ID. Cache it in Business Central to avoid repeating this.",
                    payload.Approver.Upn);

                return resolved;
            }
        }

        _logger.LogWarning(
            "No Entra object ID and no resolvable UPN for {Approver}.",
            payload.Approver.UserId);

        return null;
    }

    private async Task<(string? ConversationId, string ServiceUrl)> ResolveConversationAsync(
        string aadObjectId,
        CancellationToken cancellationToken)
    {
        // 1. Stored reference. Always preferred - it carries the service URL
        //    Teams actually used, which beats anything in configuration.
        var stored = await _conversations.GetAsync(aadObjectId, cancellationToken);

        if (stored is not null && !string.IsNullOrWhiteSpace(stored.ConversationId))
        {
            return (stored.ConversationId, stored.ServiceUrl);
        }

        // 2. Ask Teams to open one. Succeeds only if the app is installed.
        if (!_options.AttemptDirectConversation)
        {
            return (null, _options.DefaultServiceUrl);
        }

        var serviceUrl = _options.DefaultServiceUrl;
        var created = await _connector.TryCreateConversationAsync(aadObjectId, serviceUrl, cancellationToken);

        if (created is not null)
        {
            // Store it so the next dispatch skips this round trip.
            await _conversations.SaveAsync(new ConversationReference
            {
                AadObjectId = aadObjectId,
                ConversationId = created,
                ServiceUrl = serviceUrl,
                TenantId = _options.TenantId
            }, cancellationToken);
        }

        return (created, serviceUrl);
    }

    private static string BuildSummary(ApprovalDispatchPayload payload)
    {
        var party = payload.Document.CounterpartyName
                    ?? payload.Document.CounterpartyNo
                    ?? "Unknown party";

        return $"Approval needed: {payload.Document.DocumentNo} - {party} - " +
               $"{payload.Document.CurrencyCode} {payload.Document.Amount:N2}";
    }
}
