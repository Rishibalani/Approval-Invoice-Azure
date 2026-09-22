using System.Text.Json.Nodes;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Options;
using PulseNet.Approvals.Functions.Models;

namespace PulseNet.Approvals.Functions.Cards;

/// <summary>
/// Builds the Adaptive Card. One builder for every channel and every tier.
///
/// THE RULE THIS FILE EXISTS TO ENFORCE
///
/// Business Central decides whether an approval may happen in a channel. This
/// class renders that decision. It never re-derives it from the amount.
///
/// If this file ever contains a line like "if (amount > 500000)", a financial
/// control has moved out of the system of record and into a web app, and the
/// audit trail no longer tells the whole story. The amount is here to be
/// DISPLAYED, never to be judged.
///
/// Action mode, not delivery mode, decides what buttons appear:
///   NotifyOnly -> a deep link into Business Central and nothing else
///   Link       -> Action.OpenUrl carrying a signed token
///   Native     -> Action.Execute with inline comment capture, bot only
///
/// FORMATTING LIVES IN ApprovalCardViewModel, NOT HERE. This class arranges
/// strings; it does not decide how a decimal or a date looks.
/// </summary>
public sealed class ApprovalCardBuilder
{
    private readonly ChannelOptions _options;

    public ApprovalCardBuilder(IOptions<ChannelOptions> options) => _options = options.Value;

    /// <summary>
    /// 1.4, not 1.5. Teams mobile (iOS and Android) silently drops a 1.5 card
    /// sent by a bot - the message arrives as an empty space - while desktop
    /// renders it. The outcome and notice cards were always 1.4 and did show on
    /// phones; this card was the only 1.5 one.
    ///
    /// Nothing on the card needs 1.5: Input.Text label / isRequired /
    /// errorMessage are 1.3, Action.Execute and the refresh block are 1.4.
    /// </summary>
    private const string AdaptiveCardVersion = AdaptiveCardSchema.Version14;

    public JsonObject Build(
        ApprovalDispatchPayload payload,
        ChannelActionMode actionMode,
        string? approveUrl,
        string? rejectUrl)
    {
        var vm = ApprovalCardViewModel.From(payload, _options.FallbackCurrencyCode, _options.MaxLinesOnCard);

        var card = new JsonObject
        {
            ["type"] = "AdaptiveCard",
            ["$schema"] = AdaptiveCardSchema.SchemaUri,
            ["version"] = AdaptiveCardVersion,

            // WHAT A CLIENT SHOWS WHEN IT CANNOT RENDER THE CARD.
            //
            // Not optional, and its absence is invisible until somebody opens
            // Teams on a phone. A client that cannot render an Adaptive Card
            // falls back to this text - and with no fallbackText it shows
            // NOTHING AT ALL. The message appears in the conversation as an
            // empty space with a timestamp, which looks like a bug in the
            // integration rather than a rendering limit.
            //
            // Mobile Teams lags desktop on Adaptive Card support by several
            // versions, so the desktop card rendering perfectly proves very
            // little about the phone.
            //
            // The text carries the facts that matter plus where to go, so an
            // approver on an unsupported client is inconvenienced rather than
            // stuck.
            ["fallbackText"] = BuildFallbackText(vm),

            ["body"] = BuildBody(vm, payload, actionMode),
            ["actions"] = BuildActions(vm, payload, actionMode, approveUrl, rejectUrl)
        };

        // Full-width is a desktop nicety and a known source of mobile
        // rendering trouble. Channels:Teams:UseFullWidthCard - normally false;
        // turn it on only if the phones in use are known to handle it.
        if (_options.Teams.UseFullWidthCard)
        {
            card["msteams"] = new JsonObject { ["width"] = "Full" };
        }

        // Auto-refresh needs a bot to answer the invoke, and Teams ignores the
        // block when userIds is empty. Both conditions must hold.
        //
        // Also gated on configuration, because `refresh` is Adaptive Card 1.4
        // and an older mobile client that cannot parse it may drop the whole
        // card rather than just the refresh block. The card now updates itself
        // through CardRefreshService when a decision lands, so this block is a
        // secondary path rather than the only one.
        if (_options.Teams.UseCardRefreshBlock &&
            actionMode == ChannelActionMode.Native &&
            !string.IsNullOrWhiteSpace(payload.Approver.EntraObjectId))
        {
            card["refresh"] = new JsonObject
            {
                ["action"] = new JsonObject
                {
                    ["type"] = "Action.Execute",
                    ["title"] = "Refresh",
                    ["verb"] = "approval/refresh",
                    ["data"] = ActionData(payload)
                },
                ["userIds"] = new JsonArray { payload.Approver.EntraObjectId! }
            };
        }

        return card;
    }

    // ------------------------------------------------------------------
    //  Body
    // ------------------------------------------------------------------

    private static JsonArray BuildBody(
        ApprovalCardViewModel vm, ApprovalDispatchPayload payload, ChannelActionMode actionMode)
    {
        var body = new JsonArray { BuildHeader(vm) };

        if (vm.CreatedLine is not null)
        {
            body.Add(Text(vm.CreatedLine, size: "Small", subtle: true, spacing: "Small"));
        }

        if (vm.DelegatedFromName is not null)
        {
            body.Add(new JsonObject
            {
                ["type"] = "Container",
                ["style"] = "accent",
                ["bleed"] = true,
                ["spacing"] = "Small",
                ["items"] = new JsonArray
                {
                    Text($"Delegated to you by {vm.DelegatedFromName}", size: "Small", bold: true)
                }
            });
        }

        body.Add(BuildFactSet(vm));

        if (vm.Lines.Count > 0)
        {
            body.Add(BuildLinesTable(vm));
        }

        if (vm.AttachmentNote is not null)
        {
            body.Add(Text(vm.AttachmentNote, size: "Small", subtle: true, spacing: "Small"));
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
            body.Add(Text(vm.ChainContext, size: "Small", subtle: true, spacing: "Medium", separator: true));
        }

        // The comment box sits in the card body, directly above the buttons.
        //
        // It used to live inside an Action.ShowCard (a pop-out panel per
        // button with its own confirm button). Teams on iOS and Android does
        // not display a bot-sent card built that way - the message arrived as
        // an empty space - while desktop did. A flat card with one input and
        // plain Action.Execute buttons is the shape Teams mobile renders.
        //
        // Deliberately no "label", "isRequired" or "errorMessage": those make
        // the client refuse to submit Approve too when the box is empty. The
        // rejection reason is enforced by the bot instead, which answers an
        // empty Reject with a short message and leaves the card as it is.
        if (actionMode == ChannelActionMode.Native)
        {
            body.Add(Text("Comment (required to reject)", size: "Small", subtle: true, spacing: "Medium"));
            body.Add(new JsonObject
            {
                ["type"] = "Input.Text",
                ["id"] = "comment",
                ["isMultiline"] = true,
                ["maxLength"] = 250,
                ["placeholder"] = "Recorded against the approval entry in Business Central",
                ["spacing"] = "Small"
            });
        }

        return body;
    }

    private static JsonObject BuildHeader(ApprovalCardViewModel vm)
    {
        var columns = new JsonArray();

        if (vm.BrandIconUrl is not null)
        {
            columns.Add(new JsonObject
            {
                ["type"] = "Column",
                ["width"] = "auto",
                ["verticalContentAlignment"] = "Center",
                ["items"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["type"] = "Image",
                        ["url"] = vm.BrandIconUrl,
                        ["width"] = "32px",
                        ["altText"] = "Business Central"
                    }
                }
            });
        }

        columns.Add(new JsonObject
        {
            ["type"] = "Column",
            ["width"] = "stretch",
            ["verticalContentAlignment"] = "Center",
            ["items"] = new JsonArray
            {
                Text(vm.TypeCaption, size: "Small", subtle: true, spacing: "None"),
                Text($"{vm.DocumentNo} · {vm.PartyName}", size: "Medium", bold: true, spacing: "None")
            }
        });

        columns.Add(new JsonObject
        {
            ["type"] = "Column",
            ["width"] = "auto",
            ["verticalContentAlignment"] = "Center",
            ["items"] = new JsonArray
            {
                Text(vm.HeadlineAmount, size: "Large", bold: true, align: "Right", wrap: false)
            }
        });

        return new JsonObject
        {
            ["type"] = "ColumnSet",
            ["spacing"] = "None",
            ["columns"] = columns
        };
    }

    /// <summary>
    /// Facts arrive from the view model already filtered. Nothing is added
    /// here, and there is deliberately no placeholder for a missing value:
    /// FactSet values render as markdown, so a lone "-" becomes an empty
    /// bullet point. That is the floating bullet on the current card.
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
            ["spacing"] = "Medium",
            ["facts"] = facts
        };
    }

    private static JsonObject BuildLinesTable(ApprovalCardViewModel vm)
    {
        var items = new JsonArray
        {
            Text("Lines", size: "Small", bold: true),
            LineRow("Description", "Qty", "UoM", "Amount", header: true)
        };

        foreach (var line in vm.Lines)
        {
            items.Add(LineRow(line.Description, line.Quantity, line.UnitOfMeasure, line.Amount));
        }

        if (vm.HiddenLineCount > 0)
        {
            items.Add(Text(
                $"+{vm.HiddenLineCount} more line(s). Open in Business Central to see all.",
                size: "Small", subtle: true, spacing: "Small"));
        }

        return new JsonObject
        {
            ["type"] = "Container",
            ["separator"] = true,
            ["spacing"] = "Medium",
            ["items"] = items
        };
    }

    // Fixed relative column widths, identical on every row. "auto" sizes each
    // row's columns to that row's own text, so Qty / UoM / Amount drifted out
    // of line from row to row. Weighted widths are shared by all rows, so the
    // header and every line sit in the same columns on desktop and mobile.
    private const int DescriptionWeight = 44;
    private const int QtyWeight = 10;
    private const int UomWeight = 12;
    private const int AmountWeight = 34;

    private static JsonObject LineRow(string description, string qty, string uom, string amount, bool header = false) =>
        new()
        {
            ["type"] = "ColumnSet",
            ["spacing"] = "Small",
            ["separator"] = !header,
            ["columns"] = new JsonArray
            {
                LineCell(DescriptionWeight, description, header, align: "Left"),
                LineCell(QtyWeight, qty, header, align: "Right"),
                // A blank unit still needs text, or the cell collapses. A
                // non-breaking space, not " ": mobile renderers can reject a
                // TextBlock whose text is only ordinary whitespace.
                LineCell(UomWeight, string.IsNullOrEmpty(uom) ? "\u00A0" : uom, header, align: "Center"),
                LineCell(AmountWeight, amount, header, align: "Right")
            }
        };

    private static JsonObject LineCell(int weight, string text, bool header, string align) =>
        new()
        {
            ["type"] = "Column",
            ["width"] = weight,
            ["verticalContentAlignment"] = "Top",
            // wrap on: on a narrow phone a long amount wraps inside its own
            // column instead of being cut off or pushing the others.
            ["items"] = new JsonArray { Text(text, size: "Small", subtle: header, align: align, wrap: true) }
        };

    /// <summary>
    /// Renders the suppression reasons Business Central sent. The text is
    /// written for an approver, not an engineer - "VendorBankDetailsChanged"
    /// means nothing to someone holding a phone.
    /// </summary>
    private static IEnumerable<JsonObject> BuildWarnings(ApprovalDispatchPayload payload)
    {
        foreach (var reason in payload.Policy.SuppressionReasons)
        {
            var (text, colour) = reason switch
            {
                "VendorBankDetailsChanged" =>
                    ("⚠ This vendor's bank details changed after the invoice was created. Please review in Business Central before approving.", "Attention"),
                "HighValue" =>
                    ("This invoice is above the value that can be approved from a message. Please open it in Business Central.", "Warning"),
                "ChannelCeiling" =>
                    ("This invoice is above the limit for approving from this channel. Please use Business Central.", "Warning"),
                "ApproverLimitExceeded" =>
                    ("This amount is above your approval limit. Business Central will route it onward.", "Warning"),
                "DocumentChanged" =>
                    ("⚠ This invoice was edited after the request was raised. Please review it before approving.", "Attention"),
                "ApproverSuspended" =>
                    ("Channel notifications are paused for you. Please work in Business Central.", "Default"),
                _ => (string.Empty, "Default")
            };

            if (string.IsNullOrEmpty(text)) continue;

            yield return new JsonObject
            {
                ["type"] = "TextBlock",
                ["text"] = text,
                ["wrap"] = true,
                ["color"] = colour,
                ["weight"] = colour == "Attention" ? "Bolder" : "Default",
                ["spacing"] = "Medium"
            };
        }
    }

    // ------------------------------------------------------------------
    //  Actions
    // ------------------------------------------------------------------

    private static JsonArray BuildActions(
        ApprovalCardViewModel vm,
        ApprovalDispatchPayload payload,
        ChannelActionMode actionMode,
        string? approveUrl,
        string? rejectUrl)
    {
        var actions = new JsonArray();

        switch (actionMode)
        {
            case ChannelActionMode.Native:
                // Inline comment capture is only possible here: Action.Execute
                // posts the body's comment input back to the bot (Teams merges
                // it into action.data). A webhook card has nothing to post to,
                // which is why Link mode below sends the approver to a web page
                // to type a rejection reason.
                actions.Add(Execute("Approve", "approval/approve", "positive", payload));
                actions.Add(Execute("Reject", "approval/reject", "destructive", payload));

                if (vm.SubstituteName is not null && payload.Policy.CanDelegateInChannel)
                {
                    actions.Add(new JsonObject
                    {
                        ["type"] = "Action.Execute",
                        ["title"] = $"Delegate to {vm.SubstituteName}",
                        ["verb"] = "approval/delegate",
                        ["data"] = ActionData(payload)
                    });
                }
                break;

            case ChannelActionMode.Link:
                if (!string.IsNullOrWhiteSpace(approveUrl))
                {
                    actions.Add(new JsonObject
                    {
                        ["type"] = "Action.OpenUrl",
                        ["title"] = "Approve",
                        ["url"] = approveUrl,
                        ["style"] = "positive"
                    });
                }

                if (!string.IsNullOrWhiteSpace(rejectUrl))
                {
                    actions.Add(new JsonObject
                    {
                        ["type"] = "Action.OpenUrl",
                        ["title"] = "Reject",
                        ["url"] = rejectUrl,
                        ["style"] = "destructive"
                    });
                }
                break;

            case ChannelActionMode.NotifyOnly:
                // Nothing. Falls through to the deep link below.
                break;
        }

        // View in Business Central is on every card in every mode. When
        // something goes wrong - a stale token, a changed amount, a channel we
        // never anticipated - this link is the escape hatch that always works.
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

    private static JsonObject Execute(string title, string verb, string style, ApprovalDispatchPayload payload) =>
        new()
        {
            ["type"] = "Action.Execute",
            ["title"] = title,
            ["style"] = style,
            ["verb"] = verb,
            ["data"] = ActionData(payload)
        };

    /// <summary>
    /// Plain text for a client that cannot render the card.
    ///
    /// Enough to decide whether to act now or later, and where to go. No
    /// markdown - a client that cannot render a card will not render markdown
    /// either, and asterisks around a vendor name read as a fault.
    /// </summary>
    private static string BuildFallbackText(ApprovalCardViewModel vm)
    {
        var lines = new List<string>
        {
            $"{vm.TypeCaption}: {vm.DocumentNo}",
            $"{vm.PartyName} - {vm.HeadlineAmount}"
        };

        if (vm.ChainContext is not null)
        {
            lines.Add(vm.ChainContext);
        }

        lines.Add("Open the invoice in Business Central to approve or reject.");

        return string.Join("\n", lines);
    }

    private static JsonObject ActionData(ApprovalDispatchPayload payload) =>
        new()
        {
            ["approvalEntryNo"] = payload.Approval.ApprovalEntryNo,
            ["documentNo"] = payload.Document.DocumentNo,
            ["eventId"] = payload.EventId,
            ["correlationId"] = payload.CorrelationId
        };

    // ------------------------------------------------------------------

    private static JsonObject Text(
        string text,
        string? size = null,
        bool bold = false,
        bool subtle = false,
        string? spacing = null,
        string? align = null,
        bool wrap = true,
        bool separator = false)
    {
        var node = new JsonObject
        {
            ["type"] = "TextBlock",
            ["text"] = text,
            ["wrap"] = wrap
        };

        if (size is not null) node["size"] = size;
        if (bold) node["weight"] = "Bolder";
        if (subtle) node["isSubtle"] = true;
        if (spacing is not null) node["spacing"] = spacing;
        if (align is not null) node["horizontalAlignment"] = align;
        if (separator) node["separator"] = true;

        return node;
    }
}
