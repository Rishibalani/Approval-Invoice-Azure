using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Cards;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Options;

namespace PulseNet.Approvals.Functions.Channels;

/// <summary>
/// Delivers an Adaptive Card to Teams through a Power Automate webhook.
///
/// TWO PAYLOAD SHAPES, ONE SENDER
///
/// FlowRouted (the usual choice) posts { recipientUpn, summary, cardJson } to a
/// flow that reads the recipient and sends a 1:1 chat via Flow bot. That gives
/// per-approver delivery with no bot registration, no manifest and no Teams
/// admin involvement - the approver simply gets a direct message.
///
/// TeamsMessage posts the raw Teams message envelope to a channel webhook.
/// One fixed destination for everyone, because a channel webhook has its
/// target baked in at creation and cannot be told where to go. Useful for a
/// shared ops channel or a first smoke test, not for real approvals.
///
/// WHY THIS IS THE FIRST SENDER TO BUILD
///
/// It needs one URL and nothing else. No app registration, no Teams admin
/// upload, no manifest, no admin consent, no Graph permission. Paste a URL
/// into configuration and cards appear in Teams. That makes it the fastest
/// possible route to seeing real invoice data on a real card, which is the
/// thing worth proving before investing in a bot.
///
/// WHAT IT CANNOT DO, AND WHY THAT IS ACCEPTABLE
///
/// A webhook is one-way. Teams will not call us back, so Action.Execute is
/// unavailable and cards cannot refresh in place. Link mode solves this:
/// buttons are Action.OpenUrl pointing at our own action endpoint, which does
/// the work and can post a follow-up card here to show the outcome.
///
/// THE SECURITY CAVEAT THAT MATTERS
///
/// If the webhook was created from the "post to a channel" Workflows template,
/// every member of that channel sees the card and could tap Approve. The token
/// binds to the approval entry, not to whoever clicks.
///
/// That is why ActionTokenOptions.RequireSignedInUser exists. With Easy Auth
/// on, tapping triggers Entra sign-in and the action endpoint checks the
/// signed-in identity against the token's approver. Everyone can see the card;
/// only the named approver can act on it.
///
/// Without Easy Auth this is a sandbox-only mode, and the Environment Tag
/// allow-list on the ingest endpoint is what keeps it there.
/// </summary>
public sealed class WorkflowWebhookSender : IChannelSender
{
    private readonly HttpClient _http;
    private readonly ApprovalCardBuilder _cardBuilder;
    private readonly ChannelOptions _options;
    private readonly ILogger<WorkflowWebhookSender> _logger;

    public ChannelDeliveryMode Mode => ChannelDeliveryMode.WorkflowWebhook;
    public ApprovalChannel Channel => ApprovalChannel.Teams;

    public WorkflowWebhookSender(
        HttpClient http,
        ApprovalCardBuilder cardBuilder,
        IOptions<ChannelOptions> options,
        ILogger<WorkflowWebhookSender> logger)
    {
        _http = http;
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
        var webhookUrl = _options.Teams.WebhookUrl;

        if (string.IsNullOrWhiteSpace(webhookUrl))
        {
            return ChannelSendResult.Fail(Channel, "webhook_url_not_configured");
        }

        var flowRouted = _options.Teams.PayloadMode == WebhookPayloadMode.FlowRouted;

        // Routing to a person needs a person to route to. Failing here rather
        // than posting into a channel is deliberate: silently broadcasting an
        // invoice to everyone because one UPN was missing is a worse outcome
        // than not sending it.
        if (flowRouted && string.IsNullOrWhiteSpace(payload.Approver.Upn))
        {
            _logger.LogError(
                "No UPN for approver {Approver}; cannot address a 1:1 chat. " +
                "Check Authentication Email on the Business Central user.",
                payload.Approver.UserId);

            return ChannelSendResult.Fail(Channel, "no_recipient_upn");
        }

        try
        {
            var card = _cardBuilder.Build(payload, actionMode, approveUrl, rejectUrl);

            var json = flowRouted
                ? BuildFlowEnvelope(payload, card)
                : WrapForWorkflows(card).ToJsonString();

            using var content = new StringContent(json, Encoding.UTF8, "application/json");
            using var response = await _http.PostAsync(webhookUrl, content, cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                var body = await response.Content.ReadAsStringAsync(cancellationToken);

                _logger.LogError(
                    "Teams webhook rejected the card: {Status} {Body}",
                    (int)response.StatusCode, Truncate(body, 300));

                return ChannelSendResult.Fail(
                    Channel,
                    $"webhook_http_{(int)response.StatusCode}",
                    transient: (int)response.StatusCode >= 500);
            }

            _logger.LogInformation(
                "Teams card sent for {DocumentNo} to {Recipient}, mode {PayloadMode}/{ActionMode}.",
                payload.Document.DocumentNo,
                flowRouted ? payload.Approver.Upn : "(fixed channel)",
                _options.Teams.PayloadMode,
                actionMode);

            // A Workflows webhook returns no message identifier, so there is
            // nothing to correlate against for an in-place update later. That
            // capability arrives with the bot.
            return ChannelSendResult.Ok(Channel);
        }
        catch (HttpRequestException ex)
        {
            _logger.LogError(ex, "Could not reach the Teams webhook.");
            return ChannelSendResult.Fail(Channel, "webhook_unreachable", transient: true);
        }
        catch (TaskCanceledException ex)
        {
            _logger.LogError(ex, "Teams webhook timed out.");
            return ChannelSendResult.Fail(Channel, "webhook_timeout", transient: true);
        }
    }

    /// <summary>
    /// Posts a follow-up card after a decision, so the channel shows the
    /// outcome. Not an in-place update - the original card stays put with its
    /// buttons, which is the honest limitation of webhook mode. The buttons
    /// still refuse a second tap because the nonce is burned.
    /// </summary>
    public async Task PostOutcomeAsync(
        string recipientUpn,
        string documentNo,
        string outcome,
        string actedBy,
        CancellationToken cancellationToken)
    {
        var webhookUrl = _options.Teams.WebhookUrl;

        if (string.IsNullOrWhiteSpace(webhookUrl))
        {
            return;
        }

        var card = new JsonObject
        {
            ["type"] = "AdaptiveCard",
            ["$schema"] = AdaptiveCardSchema.SchemaUri,
            ["version"] = AdaptiveCardSchema.Version14,
            ["body"] = new JsonArray
            {
                new JsonObject
                {
                    ["type"] = "TextBlock",
                    ["text"] = $"{documentNo} — {outcome}",
                    ["weight"] = "Bolder",
                    ["wrap"] = true,
                    ["color"] = outcome.StartsWith("Approved", StringComparison.OrdinalIgnoreCase)
                        ? "Good"
                        : "Attention"
                },
                new JsonObject
                {
                    ["type"] = "TextBlock",
                    ["text"] = $"By {actedBy} at {DateTime.UtcNow:yyyy-MM-dd HH:mm} UTC",
                    ["isSubtle"] = true,
                    ["size"] = "Small",
                    ["spacing"] = "None",
                    ["wrap"] = true
                }
            }
        };

        try
        {
            var json = _options.Teams.PayloadMode == WebhookPayloadMode.FlowRouted
                ? new JsonObject
                {
                    ["recipientUpn"] = recipientUpn,
                    ["summary"] = $"{documentNo} {outcome}",
                    ["cardJson"] = card.ToJsonString()
                }.ToJsonString()
                : WrapForWorkflows(card).ToJsonString();

            using var content = new StringContent(json, Encoding.UTF8, "application/json");
            await _http.PostAsync(webhookUrl, content, cancellationToken);
        }
        catch (Exception ex)
        {
            // Cosmetic. The approval already succeeded; failing to announce it
            // must never surface as an error to the approver.
            _logger.LogWarning(ex, "Could not post the outcome card. The approval itself was unaffected.");
        }
    }

    /// <summary>
    /// Shape for a Power Automate routing flow.
    ///
    /// cardJson is a STRING, not a nested object. The "Post adaptive card in a
    /// chat or channel" action expects card text, and handing it an object
    /// means fiddling with string() expressions in the flow designer. Sending
    /// it pre-serialised keeps the flow to a trigger and one action.
    /// </summary>
    private static string BuildFlowEnvelope(ApprovalDispatchPayload payload, JsonObject card)
    {
        var envelope = new JsonObject
        {
            ["recipientUpn"] = payload.Approver.Upn,
            ["summary"] = BuildSummary(payload),
            ["cardJson"] = card.ToJsonString()
        };

        return envelope.ToJsonString();
    }

    /// <summary>
    /// Plain-text line for notification toasts and accessibility, where the
    /// card itself does not render.
    /// </summary>
    private static string BuildSummary(ApprovalDispatchPayload payload)
    {
        var party = payload.Document.CounterpartyName
                    ?? payload.Document.CounterpartyNo
                    ?? "Unknown party";

        return $"Approval needed: {payload.Document.DocumentNo} - {party} - " +
               $"{payload.Document.CurrencyCode} {payload.Document.Amount:N2}";
    }

    /// <summary>
    /// Workflows expects the Teams message envelope: an attachment array with
    /// the Adaptive Card content type. Posting a bare card silently produces
    /// an empty message, which is a confusing way to spend an afternoon.
    /// </summary>
    private static JsonObject WrapForWorkflows(JsonObject card) =>
        new()
        {
            ["type"] = "message",
            ["attachments"] = new JsonArray
            {
                new JsonObject
                {
                    ["contentType"] = AdaptiveCardSchema.ContentType,
                    ["contentUrl"] = null,
                    ["content"] = card.DeepClone()
                }
            }
        };

    private static string Truncate(string value, int max) =>
        value.Length <= max ? value : value[..max];
}
