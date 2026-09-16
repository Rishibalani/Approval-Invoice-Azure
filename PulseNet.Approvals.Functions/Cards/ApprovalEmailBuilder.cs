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
using System.Net;
using System.Text;
using PulseNet.Approvals.Functions.Models;

namespace PulseNet.Approvals.Functions.Cards;

/// <summary>
/// The HTML body of an approval email. Shared by both email senders.
///
/// WHY THIS IS NOT INSIDE EITHER SENDER
///
/// It was, twice. PlainEmailSender and OutlookActionableMessageSender each had
/// their own FormatMoney, Row and fact ordering, so the same invoice could
/// read differently depending on which delivery mode happened to be
/// configured - and PlainEmailSender still rendered a missing requester as a
/// bare dash, the placeholder the view model exists to eliminate.
///
/// Three copies of currency formatting is three places to fix a currency bug
/// and two places to forget.
///
/// FORMATTING LIVES IN ApprovalCardViewModel, NOT HERE. This class turns
/// already-formatted strings into table rows. If it ever needs to know what a
/// currency code means, the logic belongs in the view model instead.
///
/// WHY THE HTML LOOKS LIKE 2005
///
/// Tables and inline styles throughout, because Outlook renders HTML with
/// Word's engine: no flexbox, no grid, no external stylesheets, unreliable div
/// layout. This is what actually survives.
/// </summary>
public sealed class ApprovalEmailBuilder
{
    /// <summary>
    /// Subject line. Deliberately front-loads document number, party and
    /// amount - an approver triaging an inbox on a phone sees perhaps sixty
    /// characters, and those three are what let them decide whether to open it
    /// now or later.
    /// </summary>
    public string BuildSubject(ApprovalDispatchPayload payload)
    {
        var vm = ApprovalCardViewModel.From(payload);
        return $"Approval needed: {vm.DocumentNo} — {vm.PartyName} — {vm.HeadlineAmount}";
    }

    /// <summary>
    /// The body. actionUrls may be null for notify-only, in which case only
    /// the Business Central link is rendered.
    /// </summary>
    public string BuildBody(
        ApprovalDispatchPayload payload,
        ChannelActionMode actionMode,
        string? approveUrl,
        string? rejectUrl)
    {
        var vm = ApprovalCardViewModel.From(payload);
        var sb = new StringBuilder();

        sb.Append("<body style=\"margin:0;padding:24px;background:#f5f5f5;");
        sb.Append("font-family:Segoe UI,Helvetica,Arial,sans-serif;color:#1a1a1a;\">");
        sb.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\"><tr><td align=\"center\">");
        sb.Append("<table role=\"presentation\" width=\"580\" cellpadding=\"0\" cellspacing=\"0\" ");
        sb.Append("style=\"background:#ffffff;border-radius:8px;padding:28px;\">");

        // ---- Header ----
        sb.Append("<tr><td>");
        sb.Append($"<div style=\"color:#666;font-size:13px;\">{Enc(vm.TypeCaption)}</div>");
        sb.Append($"<div style=\"font-size:18px;font-weight:600;margin-top:2px;\">{Enc(vm.DocumentNo)} &middot; {Enc(vm.PartyName)}</div>");
        sb.Append($"<div style=\"font-size:28px;font-weight:600;margin-top:14px;\">{Enc(vm.HeadlineAmount)}</div>");

        if (vm.CreatedLine is not null)
        {
            // The view model emits markdown bold for the card. Email has no
            // markdown renderer, so it is stripped rather than shown literally.
            sb.Append($"<div style=\"color:#666;font-size:13px;margin-top:6px;\">{Enc(StripMarkdown(vm.CreatedLine))}</div>");
        }

        sb.Append("</td></tr>");

        if (vm.DelegatedFromName is not null)
        {
            sb.Append("<tr><td style=\"padding-top:14px;\">");
            sb.Append("<div style=\"background:#eff6fc;border-left:3px solid #0f6cbd;padding:10px;font-size:14px;font-weight:600;\">");
            sb.Append($"Delegated to you by {Enc(vm.DelegatedFromName)}");
            sb.Append("</div></td></tr>");
        }

        // ---- Facts. Same set, same order as the card. ----
        sb.Append("<tr><td style=\"padding-top:20px;\">");
        sb.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"font-size:14px;\">");

        foreach (var fact in vm.Facts)
        {
            sb.Append(Row(fact.Title, fact.Value));
        }

        sb.Append("</table></td></tr>");

        // ---- Lines ----
        if (vm.Lines.Count > 0)
        {
            sb.Append("<tr><td style=\"padding-top:20px;\">");
            sb.Append("<div style=\"font-size:13px;font-weight:600;margin-bottom:6px;\">Lines</div>");
            sb.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" ");
            sb.Append("style=\"font-size:13px;border-collapse:collapse;\">");
            sb.Append("<tr style=\"color:#666;\">");
            sb.Append("<td style=\"padding:4px 0;border-bottom:1px solid #e1e1e1;\">Description</td>");
            sb.Append("<td style=\"padding:4px 0;border-bottom:1px solid #e1e1e1;text-align:right;\">Qty</td>");
            sb.Append("<td style=\"padding:4px 0 4px 12px;border-bottom:1px solid #e1e1e1;text-align:right;\">Amount</td>");
            sb.Append("</tr>");

            foreach (var line in vm.Lines)
            {
                sb.Append("<tr>");
                sb.Append($"<td style=\"padding:4px 0;\">{Enc(line.Description)}</td>");
                sb.Append($"<td style=\"padding:4px 0;text-align:right;white-space:nowrap;\">{Enc(line.Quantity)}</td>");
                sb.Append($"<td style=\"padding:4px 0 4px 12px;text-align:right;white-space:nowrap;\">{Enc(line.Amount)}</td>");
                sb.Append("</tr>");
            }

            sb.Append("</table>");

            if (vm.HiddenLineCount > 0)
            {
                sb.Append($"<div style=\"color:#999;font-size:12px;margin-top:6px;\">+{vm.HiddenLineCount} more line(s). Open in Business Central to see all.</div>");
            }

            sb.Append("</td></tr>");
        }

        if (vm.AttachmentNote is not null)
        {
            sb.Append($"<tr><td style=\"padding-top:12px;color:#666;font-size:13px;\">{Enc(vm.AttachmentNote)}</td></tr>");
        }

        // ---- Warnings, above the buttons ----
        if (payload.Document.PayToDiffers)
        {
            sb.Append(Warning(
                "Payment goes to a different party than the vendor on this invoice. Worth confirming that is expected.",
                "#d97706", "#fff4e5"));
        }

        foreach (var reason in payload.Policy.SuppressionReasons)
        {
            var text = reason switch
            {
                "VendorBankDetailsChanged" => "This vendor's bank details changed after the invoice was created. Please review in Business Central before approving.",
                "HighValue" => "This invoice is above the value that can be approved from an email. Please open it in Business Central.",
                "ChannelCeiling" => "This invoice is above the limit for approving by email. Please use Business Central.",
                "ApproverLimitExceeded" => "This amount is above your approval limit. Business Central will route it onward.",
                "DocumentChanged" => "This invoice was edited after the request was raised. Please review it before approving.",
                _ => string.Empty
            };

            if (text.Length > 0)
            {
                sb.Append(Warning(text, "#a4262c", "#fdf3f4"));
            }
        }

        if (vm.ChainContext is not null)
        {
            sb.Append($"<tr><td style=\"padding-top:16px;color:#666;font-size:13px;\">{Enc(vm.ChainContext)}</td></tr>");
        }

        // ---- Buttons ----
        sb.Append("<tr><td style=\"padding-top:24px;\">");

        if (actionMode != ChannelActionMode.NotifyOnly)
        {
            if (!string.IsNullOrWhiteSpace(approveUrl))
            {
                sb.Append(Button(approveUrl, "Approve", "#107c10"));
            }

            if (!string.IsNullOrWhiteSpace(rejectUrl))
            {
                sb.Append(Button(rejectUrl, "Reject", "#a4262c"));
            }
        }

        if (!string.IsNullOrWhiteSpace(payload.Document.DeepLink))
        {
            sb.Append(Button(payload.Document.DeepLink, "View in Business Central", "#5a5a5a"));
        }

        sb.Append("</td></tr>");

        sb.Append("<tr><td style=\"padding-top:24px;color:#999;font-size:12px;\">");
        sb.Append($"Approval buttons expire {payload.Policy.ActionTokenTtlMinutes} minutes after this was sent. ");
        sb.Append("After that, please use Business Central.");
        sb.Append("</td></tr>");

        sb.Append("</table></td></tr></table></body>");

        return sb.ToString();
    }

    // ------------------------------------------------------------------

    private static string Row(string label, string value) =>
        $"<tr><td style=\"padding:4px 0;color:#666;width:150px;vertical-align:top;\">{Enc(label)}</td>" +
        $"<td style=\"padding:4px 0;\">{Enc(value)}</td></tr>";

    private static string Warning(string text, string border, string background) =>
        "<tr><td style=\"padding-top:16px;\">" +
        $"<div style=\"background:{background};border-left:3px solid {border};padding:12px;font-size:14px;\">" +
        Enc(text) + "</div></td></tr>";

    private static string Button(string url, string label, string colour) =>
        $"<a href=\"{Enc(url)}\" style=\"display:inline-block;padding:10px 20px;margin-right:8px;" +
        $"background:{colour};color:#ffffff;text-decoration:none;border-radius:4px;" +
        $"font-size:14px;font-weight:600;\">{Enc(label)}</a>";

    /// <summary>
    /// Removes the markdown bold the view model emits for Adaptive Cards.
    /// Email clients show the asterisks literally, which reads as a bug.
    /// </summary>
    private static string StripMarkdown(string value) => value.Replace("**", string.Empty);

    /// <summary>
    /// HTML-encodes every interpolated value. Vendor names come from a database
    /// somebody else can write to, and an unencoded angle bracket is how a
    /// broken layout becomes an injected link.
    /// </summary>
    private static string Enc(string? value) => WebUtility.HtmlEncode(value ?? string.Empty);
}

*/
