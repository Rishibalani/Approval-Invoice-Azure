using System.Text;
using System.Text.Json.Nodes;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Cards;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Options;

namespace PulseNet.Approvals.Functions.Channels;

/// <summary>
/// OPTION D. Hands our Adaptive Card to a Power Automate flow, which posts it
/// into the approver's 1:1 Teams chat as the Workflows (Flow) bot and waits
/// for the press.
///
/// WHY THIS EXISTS BESIDE TeamsBotSender AND WorkflowWebhookSender
///
/// TeamsBotSender (Option A) needs an Azure Bot, a Teams app package and a
/// per-approver install. It stays in place, fully configured, and comes back
/// the moment USE_POWER_AUTOMATE_CARDS is set to false.
///
/// WorkflowWebhookSender posts a card whose buttons are LINKS - the approver
/// leaves Teams for a browser. It is untouched.
///
/// This sender posts the same card with Action.Submit buttons, which the flow
/// collects. The approver stays in Teams and we still own the layout.
///
/// WHAT WE SEND, AND WHAT WE DO NOT
///
/// The envelope carries the approver's address, a one-line summary for the
/// notification toast, the card JSON and the correlation ids. It does NOT
/// carry an approval decision, an amount ceiling or anything the flow could
/// act on by itself: the flow is a delivery mechanism, not a decision maker.
/// The two signed tokens sit inside the card's own action data, so a pressed
/// button hands back a token our endpoint can verify.
///
/// THE FLOW URL IS A SECRET
///
/// A Power Automate HTTP trigger URL carries its own signature in the query
/// string. Anyone holding it can start the flow, so it belongs in Key Vault
/// and never in source control.
/// </summary>
public sealed class PowerAutomateCardSender : IChannelSender
{
    private readonly HttpClient _http;
    private readonly ApprovalCardBuilder _cardBuilder;
    private readonly ChannelOptions _options;
    private readonly ILogger<PowerAutomateCardSender> _logger;

    /// <summary>
    /// How much of a failing response to log. A display limit, not a setting:
    /// long enough to name the fault, short enough not to fill the log with a
    /// Power Automate error page.
    /// </summary>
    private const int LoggedBodyLength = 300;

    public ChannelDeliveryMode Mode => ChannelDeliveryMode.PowerAutomateCard;
    public ApprovalChannel Channel => ApprovalChannel.Teams;

    public PowerAutomateCardSender(
        HttpClient http,
        ApprovalCardBuilder cardBuilder,
        IOptions<ChannelOptions> options,
        ILogger<PowerAutomateCardSender> logger)
    {
        _http = http;
        _cardBuilder = cardBuilder;
        _options = options.Value;
        _logger = logger;
    }

    /// <param name="approveUrl">In Submit mode this is the signed Approve token, not a URL.</param>
    /// <param name="rejectUrl">In Submit mode this is the signed Reject token, not a URL.</param>
    public async Task<ChannelSendResult> SendAsync(
        ApprovalDispatchPayload payload,
        ChannelActionMode actionMode,
        string? approveUrl,
        string? rejectUrl,
        CancellationToken cancellationToken)
    {
        var flowUrl = _options.Teams.PowerAutomateFlowUrl;

        if (string.IsNullOrWhiteSpace(flowUrl))
        {
            _logger.LogError(
                "USE_POWER_AUTOMATE_CARDS is on but Channels__Teams__PowerAutomateFlowUrl is empty.");

            return ChannelSendResult.Fail(Channel, "flow_url_not_configured");
        }

        // The flow addresses a person, so it needs a person to address. Failing
        // here beats posting an invoice somewhere unintended.
        if (string.IsNullOrWhiteSpace(payload.Approver.Upn))
        {
            _logger.LogError(
                "No UPN for approver {Approver}; the flow cannot address a chat. " +
                "Check Authentication Email on the Business Central user.",
                payload.Approver.UserId);

            return ChannelSendResult.Fail(Channel, "no_recipient_upn");
        }

        try
        {
            var card = _cardBuilder.Build(payload, actionMode, approveUrl, rejectUrl);

            var envelope = new JsonObject
            {
                [FlowContract.EnvelopeRecipientUpn] = payload.Approver.Upn,
                [FlowContract.EnvelopeSummary] = BuildSummary(payload),
                [FlowContract.EnvelopeDocumentNo] = payload.Document.DocumentNo,
                [FlowContract.EnvelopeApprovalEntryNo] = payload.Approval.ApprovalEntryNo,
                [FlowContract.EnvelopeEventId] = payload.EventId,
                [FlowContract.EnvelopeCorrelationId] = payload.CorrelationId,

                // The flow posts this verbatim. Sending it as a string keeps
                // Power Automate from reshaping the card while parsing it.
                [FlowContract.EnvelopeCardJson] = card.ToJsonString(),

                // Tells the flow whether to expect a press at all. On a
                // notify-only card there is nothing to wait for.
                [FlowContract.EnvelopeExpectsResponse] = actionMode == ChannelActionMode.Submit
            };

            using var content = new StringContent(envelope.ToJsonString(), Encoding.UTF8, "application/json");
            using var response = await _http.PostAsync(flowUrl, content, cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                var body = await response.Content.ReadAsStringAsync(cancellationToken);

                _logger.LogError(
                    "The Power Automate flow rejected the card: {Status} {Body}",
                    (int)response.StatusCode, Truncate(body, LoggedBodyLength));

                return ChannelSendResult.Fail(
                    Channel,
                    $"flow_http_{(int)response.StatusCode}",
                    transient: (int)response.StatusCode >= 500);
            }

            _logger.LogInformation(
                "Card handed to Power Automate for {DocumentNo}, approver {Recipient}, mode {ActionMode}.",
                payload.Document.DocumentNo, payload.Approver.Upn, actionMode);

            // A 202 from the flow means "accepted and now waiting", not
            // "delivered". Delivery failures surface in the flow's own run
            // history, which is why the runbook has an alert step there.
            return ChannelSendResult.Ok(Channel);
        }
        catch (HttpRequestException ex)
        {
            _logger.LogError(ex, "Could not reach the Power Automate flow.");
            return ChannelSendResult.Fail(Channel, "flow_unreachable", transient: true);
        }
        catch (TaskCanceledException ex)
        {
            _logger.LogError(ex, "The Power Automate flow timed out accepting the card.");
            return ChannelSendResult.Fail(Channel, "flow_timeout", transient: true);
        }
    }

    /// <summary>
    /// Plain-text line for the notification toast and for accessibility, where
    /// the card itself does not render.
    /// </summary>
    private static string BuildSummary(ApprovalDispatchPayload payload)
    {
        var party = payload.Document.CounterpartyName
                    ?? payload.Document.CounterpartyNo
                    ?? "Unknown party";

        return $"Approval needed: {payload.Document.DocumentNo} - {party} - " +
               $"{payload.Document.CurrencyCode} {payload.Document.Amount:N2}";
    }

    private static string Truncate(string value, int max) =>
        string.IsNullOrEmpty(value) || value.Length <= max ? value : value[..max] + "...";
}
