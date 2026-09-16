// =========================================================================
//  PRESERVED - NOT IN USE
// =========================================================================
//
//  Business Central now composes and sends approval emails itself, using its
//  own email module. This file is kept rather than deleted for two reasons.
//
//  First, it still works. If Business Central email is ever unavailable in an
//  environment - or if Actionable Messages become worth their setup cost -
//  restoring it is a matter of uncommenting this file and its registration in
//  Program.cs.
//
//  Second, it documents what was tried. A reader wondering why the email path
//  moved to Business Central can see exactly what the Azure version required:
//  an app registration, admin consent, an Exchange access policy and a shared
//  mailbox, none of which Business Central needs.
//
//  WHAT DID NOT MOVE
//
//  The action endpoint. Buttons in a Business-Central-sent email still point
//  at /api/approvals/act, and Azure still validates the token, burns the
//  nonce, enforces the rejection reason and calls back. Business Central
//  composes; Azure decides.
//
//  To restore: remove this comment block and the /* */ wrapper below, then
//  uncomment the matching registration in Program.cs.
// =========================================================================

/*
using System.Text.Json.Nodes;
using PulseNet.Approvals.Functions.Models;

namespace PulseNet.Approvals.Functions.Cards;

/// <summary>
/// Builds the Adaptive Card that Outlook embeds in an email.
///
/// FORMATTING LIVES IN ApprovalCardViewModel, NOT HERE.
///
/// This class arranges elements. It does not format money, resolve a blank
/// currency code to LCY, decide which facts to show, or work out the requester
/// label. All of that is shared with the Teams card through the view model, so
/// the two cannot disagree about what an invoice says.
///
/// An earlier version of this file did its own formatting, and the result was
/// Teams showing "$10.00 CAD" while Outlook showed "CAD 10.00" for the same
/// invoice - along with missing pay-to, dimensions, lines and the requester
/// fallback chain. Every fix applied to one card then had to be remembered for
/// the other, and it was not.
///
/// WHY IT IS A SEPARATE CLASS FROM ApprovalCardBuilder
///
/// Same content, different element vocabulary. Outlook's Actionable Message
/// renderer targets Adaptive Card 1.0, and the Teams card uses several things
/// that do not exist there:
///
///   Action.Execute            no bot in email. Outlook uses Action.Http
///   Container style / bleed   1.2
///   verticalContentAlignment  1.1
///   refresh, msteams          Teams-only or later
///   action style              1.2, rendered as a default button
///   isRequired                accepted and NOT enforced. See below
///
/// Serving both from one builder means either crippling the Teams card or
/// shipping a card Outlook renders as a blank block - and it fails silently,
/// with no error in the email, the logs, or the send result.
///
/// TWO PROPERTIES OUTLOOK REQUIRES AND TEAMS DOES NOT
///
///   originator        the provider ID. Without it, no card renders at all
///   hideOriginalBody  suppresses the HTML fallback when the card does render
///
/// ON isRequired
///
/// Outlook accepts the property and ignores it. A comment box marked required
/// still submits empty, which is why the rejection-reason rule is enforced in
/// ApprovalDecisionService rather than on the card.
/// </summary>
public sealed class OutlookCardBuilder
{
    private const string AdaptiveCardVersion = "1.0";

    public JsonObject Build(
        ApprovalDispatchPayload payload,
        ChannelActionMode actionMode,
        string originatorId,
        string actionEndpointUrl,
        string? approveToken,
        string? rejectToken)
    {
        var vm = ApprovalCardViewModel.From(payload);

        var card = new JsonObject
        {
            ["type"] = "AdaptiveCard",
            ["$schema"] = "http://adaptivecards.io/schemas/adaptive-card.json",
            ["version"] = AdaptiveCardVersion,

            // Without this, Outlook silently declines to render and shows the
            // HTML fallback instead. No error appears anywhere.
            ["originator"] = originatorId,

            // The HTML body is a fallback for clients that cannot render the
            // card. When the card does render, showing both is just noise.
            ["hideOriginalBody"] = true,

            ["body"] = BuildBody(vm, payload)
        };

        var actions = BuildActions(payload, actionMode, actionEndpointUrl, approveToken, rejectToken);

        if (actions.Count > 0)
        {
            card["actions"] = actions;
        }

        return card;
    }

    // ------------------------------------------------------------------
    //  Body
    // ------------------------------------------------------------------

    private static JsonArray BuildBody(ApprovalCardViewModel vm, ApprovalDispatchPayload payload)
    {
        var body = new JsonArray();

        // Header stacked rather than the Teams ColumnSet.
        // verticalContentAlignment is 1.1, and without it the amount sits
        // awkwardly against a two-line title.
        body.Add(Text(vm.TypeCaption, size: "small", subtle: true));
        body.Add(Text($"{vm.DocumentNo} · {vm.PartyName}", size: "medium", bold: true, spacing: "none"));
        body.Add(Text(vm.HeadlineAmount, size: "large", bold: true, spacing: "small"));

        if (vm.CreatedLine is not null)
        {
            body.Add(Text(vm.CreatedLine, size: "small", subtle: true, spacing: "small"));
        }

        if (vm.DelegatedFromName is not null)
        {
            // No accent Container - style and bleed are both 1.2. Bold carries
            // the same signal at 1.0.
            body.Add(Text($"Delegated to you by {vm.DelegatedFromName}",
                size: "small", bold: true, spacing: "small"));
        }

        body.Add(BuildFactSet(vm));

        foreach (var element in BuildLines(vm))
        {
            body.Add(element);
        }

        if (vm.AttachmentNote is not null)
        {
            body.Add(Text(vm.AttachmentNote, size: "small", subtle: true, spacing: "small"));
        }

        // Warnings before the buttons, always. An approver who has already
        // decided by the time they scroll past the amount will not read a
        // caveat placed underneath it.
        foreach (var warning in BuildWarnings(payload))
        {
            body.Add(warning);
        }

        if (vm.ChainContext is not null)
        {
            body.Add(Text(vm.ChainContext, size: "small", subtle: true, spacing: "medium", separator: true));
        }

        return body;
    }

    /// <summary>
    /// Facts arrive from the view model already filtered. Nothing is added
    /// here, and there is deliberately no placeholder for a missing value -
    /// FactSet values render as markdown, so a lone dash becomes an empty
    /// bullet point.
    /// </summary>
    private static JsonObject BuildFactSet(ApprovalCardViewModel vm)
    {
        var facts = new JsonArray();

        foreach (var fact in vm.Facts)
        {
            facts.Add(new JsonObject
            {
                ["title"] = fact.Title,
                ["value"] = fact.Value
            });
        }

        return new JsonObject
        {
            ["type"] = "FactSet",
            ["separator"] = true,
            ["facts"] = facts
        };
    }

    /// <summary>
    /// Lines as a second FactSet rather than the Teams ColumnSet grid.
    ///
    /// ColumnSet exists in 1.0, but "stretch" and "auto" column widths behave
    /// inconsistently across Outlook clients, and a misaligned three-column
    /// table in a narrow reading pane is worse than a plain list. Description
    /// as the title, quantity and amount as the value, stays legible
    /// everywhere.
    /// </summary>
    private static IEnumerable<JsonObject> BuildLines(ApprovalCardViewModel vm)
    {
        if (vm.Lines.Count == 0)
        {
            yield break;
        }

        yield return Text("Lines", size: "small", bold: true, spacing: "medium", separator: true);

        var facts = new JsonArray();

        foreach (var line in vm.Lines)
        {
            facts.Add(new JsonObject
            {
                ["title"] = line.Description,
                ["value"] = $"{line.Quantity}  ·  {line.Amount}"
            });
        }

        yield return new JsonObject
        {
            ["type"] = "FactSet",
            ["facts"] = facts
        };

        if (vm.HiddenLineCount > 0)
        {
            yield return Text(
                $"+{vm.HiddenLineCount} more line(s). Open in Business Central to see all.",
                size: "small", subtle: true, spacing: "small");
        }
    }

    private static IEnumerable<JsonObject> BuildWarnings(ApprovalDispatchPayload payload)
    {
        if (payload.Document.PayToDiffers)
        {
            yield return Text(
                "Payment goes to a different party than the vendor on this invoice. Worth confirming that is expected.",
                colour: "warning", spacing: "medium");
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

            yield return Text(text, colour: colour, bold: colour == "attention", spacing: "medium");
        }
    }

    // ------------------------------------------------------------------
    //  Actions
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
                actions.Add(DecisionAction("Approve", "approve", actionEndpointUrl, approveToken, reasonRequired: false));
            }

            if (!string.IsNullOrWhiteSpace(rejectToken))
            {
                actions.Add(DecisionAction("Reject", "reject", actionEndpointUrl, rejectToken, reasonRequired: true));
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
    /// A decision button, wrapped in ShowCard so the approver can add a comment
    /// before committing.
    ///
    /// The wrapper is not decoration. A bare Action.Http fires the instant it
    /// is clicked, and in an inbox a stray click while scrolling is entirely
    /// plausible - too easy for something that releases a payment. ShowCard
    /// makes approving deliberate and gives somewhere to record why.
    ///
    /// On Reject the placeholder says the reason is required, because
    /// ApprovalDecisionService refuses an empty one. That refusal is the actual
    /// enforcement; this text only stops the approver discovering it by being
    /// turned away.
    /// </summary>
    private static JsonObject DecisionAction(
        string title,
        string verb,
        string actionEndpointUrl,
        string token,
        bool reasonRequired)
    {
        // {{comment.value}} is substituted by Outlook from the input below. The
        // body is a STRING, not an object - Outlook posts it verbatim, so the
        // inner quotes are escaped.
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
                        ["placeholder"] = reasonRequired
                            ? "Why are you rejecting this? (required)"
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

    /// <summary>
    /// A TextBlock. Property values are lower-case because Outlook's 1.0
    /// renderer is stricter about casing than the Teams renderer, and a
    /// capitalised size is ignored rather than rejected - which looks like a
    /// styling bug rather than a schema one.
    /// </summary>
    private static JsonObject Text(
        string text,
        string? size = null,
        bool bold = false,
        bool subtle = false,
        string? colour = null,
        string? spacing = null,
        bool separator = false,
        bool wrap = true)
    {
        var block = new JsonObject
        {
            ["type"] = "TextBlock",
            ["text"] = text,
            ["wrap"] = wrap
        };

        if (size is not null) block["size"] = size;
        if (bold) block["weight"] = "bolder";
        if (subtle) block["isSubtle"] = true;
        if (colour is not null) block["color"] = colour;
        if (spacing is not null) block["spacing"] = spacing;
        if (separator) block["separator"] = true;

        return block;
    }
}

*/
