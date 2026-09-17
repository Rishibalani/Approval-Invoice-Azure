using System.Text.Json.Nodes;
using Microsoft.Extensions.Logging;
using PulseNet.Approvals.Functions.Models;

namespace PulseNet.Approvals.Functions.Services;

/// <summary>
/// Replaces a Teams card with the outcome once an approval is decided.
///
/// WHY THIS EXISTS
///
/// A decision can be made in four places: the Teams card, the Outlook email,
/// a WhatsApp message, or Business Central itself. Whichever is used, the
/// others are left showing a card that looks like it is still waiting.
///
/// The buttons on those stale cards already fail safely - Business Central
/// refuses a second decision and the approver gets "already handled". But an
/// approver seeing an Approve button has no way to know it will not work, and
/// pressing it to find out is a poor experience for something that should
/// simply have updated itself.
///
/// WHAT CAN AND CANNOT BE REFRESHED
///
///   Teams     yes. A bot can replace a message it sent
///   Outlook   NO. An email that has left cannot be changed. The only option
///             would be a second email, which means two per approval - worse
///             than a stale one
///   WhatsApp  NO. Same reason
///
/// So this is Teams-only, and deliberately so rather than by oversight.
/// </summary>
public sealed class CardRefreshService
{
    private readonly SentCardStore _sentCards;
    private readonly BotConnectorClient _connector;
    private readonly ILogger<CardRefreshService> _logger;

    public CardRefreshService(
        SentCardStore sentCards,
        BotConnectorClient connector,
        ILogger<CardRefreshService> logger)
    {
        _sentCards = sentCards;
        _connector = connector;
        _logger = logger;
    }

    /// <summary>
    /// Refreshes the card for one approval entry, if one was sent.
    ///
    /// Silent when there is nothing to refresh. An approval decided before a
    /// card went out, or one whose channel was Outlook only, is a normal case
    /// and not worth a warning.
    /// </summary>
    public async Task RefreshAsync(
        ApprovalDispatchPayload payload,
        CancellationToken cancellationToken)
    {
        var entryNo = payload.Approval.ApprovalEntryNo;
        var card = await _sentCards.GetAsync(entryNo, cancellationToken);

        if (card is null)
        {
            _logger.LogInformation(
                "No Teams card recorded for entry {EntryNo}; nothing to refresh.", entryNo);
            return;
        }

        var outcome = DescribeOutcome(payload);

        var replacement = BuildOutcomeCard(
            card.DocumentNo.Length > 0 ? card.DocumentNo : payload.Document.DocumentNo,
            outcome.Title,
            outcome.Detail,
            outcome.IsPositive);

        var updated = await _connector.UpdateCardAsync(
            card.ConversationId,
            card.ActivityId,
            card.ServiceUrl,
            replacement,
            $"{payload.Document.DocumentNo} {outcome.Title}",
            cancellationToken);

        if (updated)
        {
            _logger.LogInformation(
                "Teams card for entry {EntryNo} updated to {Outcome}.", entryNo, outcome.Title);

            // Cleared so a later event for the same entry does not try to
            // update a card that already shows an outcome.
            await _sentCards.DeleteAsync(entryNo, cancellationToken);
        }
        else
        {
            // Left in place deliberately. The update may have failed for a
            // transient reason, and a retry costs nothing.
            _logger.LogWarning(
                "Could not update the Teams card for entry {EntryNo}. " +
                "The decision is recorded in Business Central regardless.", entryNo);
        }
    }

    private static (string Title, string Detail, bool IsPositive) DescribeOutcome(
        ApprovalDispatchPayload payload)
    {
        // The event type is the wire name, not the caption - "Approved", not
        // "Approval Approved".
        var actor = payload.Approval.RequesterLabel;

        return payload.EventType switch
        {
            "Approved" => (
                "Approved",
                "This invoice has been approved. No further action is needed here.",
                true),

            "Rejected" => (
                "Rejected",
                "This invoice has been rejected and returned to the requester.",
                false),

            "Cancelled" => (
                "Cancelled",
                "The approval request was cancelled before a decision was made.",
                false),

            _ => (
                "Handled",
                "This request has already been dealt with.",
                true)
        };
    }

    private static JsonObject BuildOutcomeCard(
        string documentNo, string title, string detail, bool isPositive) =>
        new()
        {
            ["type"] = "AdaptiveCard",
            ["$schema"] = "http://adaptivecards.io/schemas/adaptive-card.json",
            ["version"] = "1.4",
            ["body"] = new JsonArray
            {
                new JsonObject
                {
                    ["type"] = "TextBlock",
                    ["text"] = $"{documentNo} — {title}",
                    ["weight"] = "Bolder",
                    ["size"] = "Medium",
                    ["color"] = isPositive ? "Good" : "Attention",
                    ["wrap"] = true
                },
                new JsonObject
                {
                    ["type"] = "TextBlock",
                    ["text"] = detail,
                    ["wrap"] = true,
                    ["isSubtle"] = true,
                    ["spacing"] = "Small"
                },
                new JsonObject
                {
                    // Deliberately not "by whom". The status event carries who
                    // the approval was assigned to, not who actually pressed
                    // the button - and on a chain those differ. Naming the
                    // wrong person in an outcome message is worse than naming
                    // nobody; Business Central's audit trail has the truth.
                    ["type"] = "TextBlock",
                    ["text"] = $"Updated {DateTime.UtcNow:MM/dd/yy HH:mm} UTC",
                    ["size"] = "Small",
                    ["isSubtle"] = true,
                    ["spacing"] = "Small"
                }
            }
        };
}
