using System.Text.Json.Nodes;
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
    private const string AdaptiveCardSchema = "http://adaptivecards.io/schemas/adaptive-card.json";

    /// <summary>
    /// 1.5 unlocks Input.Text validation and the refresh block. Teams renders
    /// it and the Workflows webhook path renders it. Outlook actionable
    /// messages do NOT - that channel is capped at 1.0, which is why
    /// PlainEmailSender builds its own simpler body.
    /// </summary>
    private const string AdaptiveCardVersion = "1.5";

    public JsonObject Build(
        ApprovalDispatchPayload payload,
        ChannelActionMode actionMode,
        string? approveUrl,
        string? rejectUrl)
    {
        var vm = ApprovalCardViewModel.From(payload);

        var card = new JsonObject
        {
            ["type"] = "AdaptiveCard",
            ["$schema"] = AdaptiveCardSchema,
            ["version"] = AdaptiveCardVersion,
            ["msteams"] = new JsonObject { ["width"] = "Full" },
            ["body"] = BuildBody(vm, payload),
            ["actions"] = BuildActions(vm, payload, actionMode, approveUrl, rejectUrl)
        };

        // Auto-refresh needs a bot to answer the invoke, and Teams ignores the
        // block when userIds is empty. Both conditions must hold.
        if (actionMode == ChannelActionMode.Native &&
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

    private static JsonArray BuildBody(ApprovalCardViewModel vm, ApprovalDispatchPayload payload)
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
            LineRow("Description", "Qty", "Amount", header: true)
        };

        foreach (var line in vm.Lines)
        {
            items.Add(LineRow(line.Description, line.Quantity, line.Amount));
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

    private static JsonObject LineRow(string description, string qty, string amount, bool header = false) =>
        new()
        {
            ["type"] = "ColumnSet",
            ["spacing"] = "Small",
            ["separator"] = !header,
            ["columns"] = new JsonArray
            {
                new JsonObject
                {
                    ["type"] = "Column",
                    ["width"] = "stretch",
                    ["items"] = new JsonArray { Text(description, size: "Small", subtle: header) }
                },
                new JsonObject
                {
                    ["type"] = "Column",
                    ["width"] = "auto",
                    ["items"] = new JsonArray { Text(qty, size: "Small", subtle: header, align: "Right", wrap: false) }
                },
                new JsonObject
                {
                    ["type"] = "Column",
                    ["width"] = "auto",
                    ["items"] = new JsonArray { Text(amount, size: "Small", subtle: header, align: "Right", wrap: false) }
                }
            }
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
                // Inline comment capture is only possible here. Action.Execute
                // posts the ShowCard's inputs back to the bot. A webhook card
                // has nothing to post to, which is why Link mode below sends
                // the approver to a web page to type a rejection reason.
                actions.Add(ExecuteWithComment(
                    title: "Approve",
                    verb: "approval/approve",
                    style: "positive",
                    label: "Comment (optional)",
                    required: false,
                    payload: payload));

                actions.Add(ExecuteWithComment(
                    title: "Reject",
                    verb: "approval/reject",
                    style: "destructive",
                    label: "Reason for rejection",
                    required: true,
                    payload: payload));

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

    private static JsonObject ExecuteWithComment(
        string title, string verb, string style, string label, bool required,
        ApprovalDispatchPayload payload)
    {
        var input = new JsonObject
        {
            ["type"] = "Input.Text",
            ["id"] = "comment",
            ["label"] = label,
            ["isMultiline"] = true,
            ["maxLength"] = 250
        };

        if (required)
        {
            // Client-side gate only. The bot re-checks server-side; an
            // Adaptive Card input is a convenience, never a control.
            input["isRequired"] = true;
            input["errorMessage"] = "A rejection reason is required.";
        }
        else
        {
            input["placeholder"] = "Recorded against the approval entry in Business Central";
        }

        return new JsonObject
        {
            ["type"] = "Action.ShowCard",
            ["title"] = title,
            ["style"] = style,
            ["card"] = new JsonObject
            {
                ["type"] = "AdaptiveCard",
                ["version"] = AdaptiveCardVersion,
                ["body"] = new JsonArray { input },
                ["actions"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["type"] = "Action.Execute",
                        ["title"] = required ? "Confirm rejection" : "Confirm approval",
                        ["style"] = style,
                        ["verb"] = verb,
                        ["data"] = ActionData(payload)
                    }
                }
            }
        };
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
