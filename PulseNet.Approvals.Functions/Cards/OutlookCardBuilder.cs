using System.Globalization;
using System.Text.Json.Nodes;
using PulseNet.Approvals.Functions.Models;

namespace PulseNet.Approvals.Functions.Cards;

/// <summary>
/// Builds the Adaptive Card that Outlook embeds in an email.
///
/// WHY THIS IS NOT ApprovalCardBuilder
///
/// Same fields, different vocabulary. Outlook's Actionable Message renderer
/// targets Adaptive Card 1.0, and the Teams card uses several things that do
/// not exist there:
///
///   Action.Execute      no bot, no invoke channel - Outlook uses Action.Http
///   refresh             1.4, ignored
///   msteams             Teams-only, ignored
///   action style        1.2. Outlook renders default buttons regardless
///   isRequired          declared but NOT enforced client-side. See below
///
/// Trying to serve both from one builder means either crippling the Teams card
/// or shipping a card Outlook renders as a blank block, which is how it fails -
/// no error, no fallback, just nothing where the card should be.
///
/// TWO PROPERTIES OUTLOOK REQUIRES AND TEAMS DOES NOT
///
///   originator  the provider ID from the Actionable Email Developer
///               Dashboard. Without it Outlook silently refuses to render.
///   hideOriginalBody  suppresses the HTML fallback when the card renders,
///               so the recipient does not see the same content twice.
///
/// ON isRequired
///
/// Outlook accepts the property and does not enforce it. A comment box marked
/// required can still be submitted empty. Any input that matters must be
/// validated server-side, which is where it should have been anyway.
/// </summary>
public sealed class OutlookCardBuilder
{
    private const string AdaptiveCardVersion = "1.0";

    /// <summary>
    /// Builds the card. Action URLs point at the Outlook action endpoint and
    /// carry the signed token in the POST body rather than the query string -
    /// a body is not written to browser history, proxy logs or referrer
    /// headers the way a URL is.
    /// </summary>
    public JsonObject Build(
        ApprovalDispatchPayload payload,
        ChannelActionMode actionMode,
        string originatorId,
        string actionEndpointUrl,
        string? approveToken,
        string? rejectToken)
    {
        var card = new JsonObject
        {
            ["type"] = "AdaptiveCard",
            ["$schema"] = "http://adaptivecards.io/schemas/adaptive-card.json",
            ["version"] = AdaptiveCardVersion,

            // Without this, Outlook renders nothing and reports nothing.
            ["originator"] = originatorId,

            // The HTML body is a fallback for clients that cannot render the
            // card. When the card does render, showing both is just noise.
            ["hideOriginalBody"] = true,

            ["body"] = BuildBody(payload)
        };

        var actions = BuildActions(payload, actionMode, actionEndpointUrl, approveToken, rejectToken);

        if (actions.Count > 0)
        {
            card["actions"] = actions;
        }

        return card;
    }

    // ------------------------------------------------------------------
    //  Body - same information as the Teams card, 1.0 elements only
    // ------------------------------------------------------------------

    private static JsonArray BuildBody(ApprovalDispatchPayload payload)
    {
        var isPayable = payload.Document.Direction == "Payable";

        var body = new JsonArray
        {
            new JsonObject
            {
                ["type"] = "TextBlock",
                ["text"] = isPayable ? "Purchase invoice approval" : "Sales invoice approval",
                ["weight"] = "bolder",
                ["size"] = "medium",
                ["wrap"] = true
            },
            new JsonObject
            {
                ["type"] = "TextBlock",
                ["text"] = payload.Document.CounterpartyName
                           ?? payload.Document.CounterpartyNo
                           ?? "Unknown party",
                ["isSubtle"] = true,
                ["spacing"] = "none",
                ["wrap"] = true
            },
            new JsonObject
            {
                ["type"] = "TextBlock",
                // Tax-inclusive: the figure that leaves the bank account, and
                // the number an approver reads first.
                ["text"] = FormatMoney(
                    payload.Document.AmountInclTax != 0
                        ? payload.Document.AmountInclTax
                        : payload.Document.Amount,
                    payload.Document.CurrencyCode),
                ["size"] = "large",
                ["weight"] = "bolder",
                ["spacing"] = "small"
            },
            BuildFactSet(payload)
        };

        foreach (var warning in BuildWarnings(payload))
        {
            body.Add(warning);
        }

        if (payload.Approval.TotalStepsInChain > 1)
        {
            body.Add(new JsonObject
            {
                ["type"] = "TextBlock",
                ["text"] = payload.Approval.IsFinalStep
                    ? $"Approval {payload.Approval.SequenceNo} of {payload.Approval.TotalStepsInChain}. This is the final approval — approving releases the invoice."
                    : $"Approval {payload.Approval.SequenceNo} of {payload.Approval.TotalStepsInChain}. Further approval is required after yours.",
                ["isSubtle"] = true,
                ["size"] = "small",
                ["wrap"] = true,
                ["spacing"] = "medium"
            });
        }

        return body;
    }

    /// <summary>
    /// Empty values are omitted rather than rendered as a dash. A row reading
    /// "Due: -" looks like a fault and costs a line on a card that is already
    /// competing with an inbox for attention.
    /// </summary>
    private static JsonObject BuildFactSet(ApprovalDispatchPayload payload)
    {
        var facts = new JsonArray();

        AddFact(facts, "Document", payload.Document.DocumentNo);
        AddFact(facts, "Their reference", payload.Document.ExternalDocumentNo);

        AddFact(facts,
            payload.Document.Direction == "Payable" ? "Vendor" : "Customer",
            ComposeParty(payload.Document.CounterpartyName, payload.Document.CounterpartyNo));

        if (payload.Document.PayToDiffers)
        {
            AddFact(facts, "Pay-to", ComposeParty(payload.Document.PayToName, payload.Document.PayToNo));
        }

        var currency = payload.Document.CurrencyCode;

        AddFact(facts, "Amount excl. tax", FormatMoney(payload.Document.AmountExclTax, currency));

        if (payload.Document.AmountInclTax != payload.Document.AmountExclTax)
        {
            AddFact(facts, "Amount incl. tax", FormatMoney(payload.Document.AmountInclTax, currency));
        }

        if (!string.IsNullOrWhiteSpace(currency) &&
            payload.Document.Amount != payload.Document.AmountLcy)
        {
            AddFact(facts, "Local value", FormatMoney(payload.Document.AmountLcy, null));
        }

        AddFact(facts, "Document date", payload.Document.DocumentDate);
        AddFact(facts, "Posting date", payload.Document.PostingDate);
        AddFact(facts, "Due", payload.Document.DueDate);
        AddFact(facts, "Requested by", payload.Approval.RequesterLabel);

        return new JsonObject
        {
            ["type"] = "FactSet",
            ["separator"] = true,
            ["facts"] = facts
        };
    }

    private static IEnumerable<JsonObject> BuildWarnings(ApprovalDispatchPayload payload)
    {
        if (payload.Document.PayToDiffers)
        {
            yield return new JsonObject
            {
                ["type"] = "TextBlock",
                ["text"] = "Payment goes to a different party than the vendor on this invoice. Worth confirming that is expected.",
                ["wrap"] = true,
                ["color"] = "warning",
                ["spacing"] = "medium"
            };
        }

        foreach (var reason in payload.Policy.SuppressionReasons)
        {
            var (text, colour) = reason switch
            {
                "VendorBankDetailsChanged" =>
                    ("This vendor's bank details changed after the invoice was created. Please review in Business Central before approving.", "attention"),
                "HighValue" =>
                    ("This invoice is above the value that can be approved from an email. Please open it in Business Central.", "warning"),
                "ChannelCeiling" =>
                    ("This invoice is above the limit for approving by email. Please use Business Central.", "warning"),
                "ApproverLimitExceeded" =>
                    ("This amount is above your approval limit. Business Central will route it onward.", "warning"),
                "DocumentChanged" =>
                    ("This invoice was edited after the request was raised. Please review it before approving.", "attention"),
                _ => (string.Empty, "default")
            };

            if (text.Length == 0)
            {
                continue;
            }

            yield return new JsonObject
            {
                ["type"] = "TextBlock",
                ["text"] = text,
                ["wrap"] = true,
                ["color"] = colour,
                ["weight"] = colour == "attention" ? "bolder" : "default",
                ["spacing"] = "medium"
            };
        }
    }

    // ------------------------------------------------------------------
    //  Actions - Action.Http, wrapped in Action.ShowCard for a comment
    // ------------------------------------------------------------------

    private static JsonArray BuildActions(
        ApprovalDispatchPayload payload,
        ChannelActionMode actionMode,
        string actionEndpointUrl,
        string? approveToken,
        string? rejectToken)
    {
        var actions = new JsonArray();

        if (actionMode != ChannelActionMode.NotifyOnly)
        {
            if (!string.IsNullOrWhiteSpace(approveToken))
            {
                actions.Add(DecisionAction("Approve", "approve", actionEndpointUrl, approveToken));
            }

            if (!string.IsNullOrWhiteSpace(rejectToken))
            {
                actions.Add(DecisionAction("Reject", "reject", actionEndpointUrl, rejectToken));
            }
        }

        // On every card in every mode. When something goes wrong - a stale
        // token, a changed amount, a client that renders nothing - this is the
        // escape hatch that always works.
        if (!string.IsNullOrWhiteSpace(payload.Document.DeepLink))
        {
            actions.Add(new JsonObject
            {
                ["type"] = "Action.OpenUrl",
                ["title"] = "View in Business Central",
                ["url"] = payload.Document.DeepLink
            });
        }

        return actions;
    }

    /// <summary>
    /// A decision button, wrapped in ShowCard so the approver can add a
    /// comment before committing.
    ///
    /// The wrapper is not decoration. A bare Action.Http fires the instant it
    /// is clicked, with no confirmation step - in an inbox, where a stray click
    /// while scrolling is entirely plausible, that is too easy. ShowCard turns
    /// approving into a deliberate two-step act and gives somewhere to record
    /// why.
    ///
    /// The comment is optional and is NOT enforced by Outlook: isRequired is
    /// accepted and ignored. Validate anything that matters server-side.
    /// </summary>
    private static JsonObject DecisionAction(
        string title,
        string verb,
        string actionEndpointUrl,
        string token)
    {
        // {{comment.value}} is substituted by Outlook from the input below.
        // Serialised carefully because the body is a STRING, not an object -
        // Outlook posts it verbatim.
        var bodyTemplate =
            $"{{\"token\":\"{token}\",\"verb\":\"{verb}\",\"comment\":\"{{{{comment.value}}}}\"}}";

        return new JsonObject
        {
            ["type"] = "Action.ShowCard",
            ["title"] = title,
            ["card"] = new JsonObject
            {
                ["type"] = "AdaptiveCard",
                ["version"] = AdaptiveCardVersion,
                ["body"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["type"] = "Input.Text",
                        ["id"] = "comment",
                        ["isMultiline"] = true,
                        ["placeholder"] = verb == "reject"
                            ? "Why are you rejecting this? (optional)"
                            : "Add a comment (optional)"
                    }
                },
                ["actions"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["type"] = "Action.Http",
                        ["title"] = $"Confirm {title.ToLowerInvariant()}",
                        ["method"] = "POST",
                        ["url"] = actionEndpointUrl,
                        ["body"] = bodyTemplate,
                        ["headers"] = new JsonArray
                        {
                            new JsonObject
                            {
                                ["name"] = "Content-Type",
                                ["value"] = "application/json"
                            }
                        }
                    }
                }
            }
        };
    }

    // ------------------------------------------------------------------

    private static void AddFact(JsonArray facts, string title, string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || value == "-")
        {
            return;
        }

        facts.Add(new JsonObject { ["title"] = title, ["value"] = value });
    }

    private static string ComposeParty(string? name, string? number)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return number ?? string.Empty;
        }

        return string.IsNullOrWhiteSpace(number) ? name : $"{name} ({number})";
    }

    private static string FormatMoney(decimal amount, string? currencyCode)
    {
        var formatted = amount.ToString("N2", CultureInfo.InvariantCulture);
        return string.IsNullOrWhiteSpace(currencyCode) ? formatted : $"{currencyCode} {formatted}";
    }
}
