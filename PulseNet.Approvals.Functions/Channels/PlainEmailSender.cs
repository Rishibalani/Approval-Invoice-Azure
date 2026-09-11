using System.Globalization;
using System.Net;
using System.Text;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Options;

namespace PulseNet.Approvals.Functions.Channels;

/// <summary>
/// Plain HTML email with link buttons. The universal fallback.
///
/// Every approver has a mailbox. Not every approver has the Teams app
/// installed, a verified WhatsApp number, or a working Actionable Messages
/// registration. This is the channel that always resolves to something.
///
/// TRANSPORT IS DELIBERATELY NOT WIRED
///
/// The HTML is complete and correct. What is missing is the thing that puts
/// it on the wire, because that is a real decision with cost and permission
/// consequences and it has not been made yet:
///
///   1. Microsoft Graph sendMail
///      Needs Mail.Send application permission plus admin consent, and a
///      mailbox to send as. Mail lands in that mailbox's Sent Items, which
///      auditors tend to like.
///
///   2. Azure Communication Services Email
///      No Graph permission, no admin consent. Needs a verified domain and
///      costs per message. Mail comes from a service domain unless you verify
///      your own, so it can look less trustworthy to recipients.
///
///   3. Send from Business Central instead
///      BC already has a working email setup in your tenant. The dispatcher
///      would return "email me this one" rather than sending it. Zero new
///      infrastructure; the cost is a round trip and a more tangled flow.
///
/// Until one is chosen, SendAsync logs the full HTML at Information level.
/// That makes the card content reviewable end to end without committing to a
/// transport, and it is deliberately loud rather than silently succeeding -
/// a sender that pretends to work is worse than one that admits it cannot.
/// </summary>
public sealed class PlainEmailSender : IChannelSender
{
    private readonly ChannelOptions _options;
    private readonly ILogger<PlainEmailSender> _logger;

    public ChannelDeliveryMode Mode => ChannelDeliveryMode.PlainEmail;
    public ApprovalChannel Channel => ApprovalChannel.Outlook;

    public PlainEmailSender(
        IOptions<ChannelOptions> options,
        ILogger<PlainEmailSender> logger)
    {
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
        var recipient = payload.Approver.Upn;

        if (string.IsNullOrWhiteSpace(recipient))
        {
            // Without a UPN there is no mailbox, and this is the fallback -
            // so nothing downstream can rescue it. Worth an error.
            _logger.LogError(
                "No UPN for approver {Approver}; the email fallback cannot deliver.",
                payload.Approver.UserId);

            return ChannelSendResult.Fail(Channel, "no_recipient_address");
        }

        var subject = BuildSubject(payload);
        var html = BuildHtml(payload, actionMode, approveUrl, rejectUrl);

        switch (_options.Outlook.Transport.ToLowerInvariant())
        {
            case "graph":
                return ChannelSendResult.Fail(Channel, "transport_graph_not_implemented");

            case "acs":
                return ChannelSendResult.Fail(Channel, "transport_acs_not_implemented");

            default:
                _logger.LogInformation(
                    "EMAIL NOT SENT - no transport configured.\nTo: {To}\nSubject: {Subject}\n{Html}",
                    recipient, subject, html);

                await Task.CompletedTask;

                return ChannelSendResult.Fail(
                    Channel,
                    "email_transport_not_configured",
                    transient: false);
        }
    }

    private static string BuildSubject(ApprovalDispatchPayload payload)
    {
        var party = payload.Document.CounterpartyName ?? payload.Document.CounterpartyNo ?? "Unknown";
        var amount = FormatMoney(payload.Document.Amount, payload.Document.CurrencyCode);

        return $"Approval needed: {payload.Document.DocumentNo} — {party} — {amount}";
    }

    /// <summary>
    /// Table-based layout with inline styles throughout.
    ///
    /// This looks like 2005 web development because email clients are 2005 web
    /// browsers. Outlook renders HTML with Microsoft Word's engine: no flexbox,
    /// no grid, no external stylesheets, and unreliable support for div-based
    /// layout. Tables and inline styles are what actually survive.
    /// </summary>
    private string BuildHtml(
        ApprovalDispatchPayload payload,
        ChannelActionMode actionMode,
        string? approveUrl,
        string? rejectUrl)
    {
        var sb = new StringBuilder();

        var isPayable = payload.Document.Direction == "Payable";
        var party = Encode(payload.Document.CounterpartyName ?? payload.Document.CounterpartyNo ?? "Unknown party");
        var amount = Encode(FormatMoney(payload.Document.Amount, payload.Document.CurrencyCode));

        sb.Append("<html><body style=\"margin:0;padding:24px;background:#f5f5f5;");
        sb.Append("font-family:Segoe UI,Helvetica,Arial,sans-serif;color:#1a1a1a;\">");
        sb.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\"><tr><td align=\"center\">");
        sb.Append("<table role=\"presentation\" width=\"560\" cellpadding=\"0\" cellspacing=\"0\" ");
        sb.Append("style=\"background:#ffffff;border-radius:8px;padding:28px;\">");

        // Header
        sb.Append("<tr><td>");
        sb.Append($"<div style=\"font-size:18px;font-weight:600;\">{(isPayable ? "Purchase invoice approval" : "Sales invoice approval")}</div>");
        sb.Append($"<div style=\"color:#666;margin-top:4px;\">{party}</div>");
        sb.Append($"<div style=\"font-size:28px;font-weight:600;margin-top:16px;\">{amount}</div>");
        sb.Append("</td></tr>");

        // Facts
        sb.Append("<tr><td style=\"padding-top:20px;\">");
        sb.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"font-size:14px;\">");
        sb.Append(Row("Document", payload.Document.DocumentNo));

        if (!string.IsNullOrWhiteSpace(payload.Document.ExternalDocumentNo))
        {
            sb.Append(Row("Their reference", payload.Document.ExternalDocumentNo));
        }

        if (!string.IsNullOrWhiteSpace(payload.Document.DueDate))
        {
            sb.Append(Row("Due", payload.Document.DueDate));
        }

        sb.Append(Row("Requested by", payload.Approval.RequestedBy ?? "-"));
        sb.Append("</table></td></tr>");

        // Warnings, before the buttons
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
                sb.Append("<tr><td style=\"padding-top:16px;\">");
                sb.Append("<div style=\"background:#fff4e5;border-left:3px solid #d97706;padding:12px;font-size:14px;\">");
                sb.Append(Encode(text));
                sb.Append("</div></td></tr>");
            }
        }

        // Chain context
        if (payload.Approval.TotalStepsInChain > 1)
        {
            var chainText = payload.Approval.IsFinalStep
                ? $"Approval {payload.Approval.SequenceNo} of {payload.Approval.TotalStepsInChain}. This is the final approval — approving releases the invoice."
                : $"Approval {payload.Approval.SequenceNo} of {payload.Approval.TotalStepsInChain}. Further approval is required after yours.";

            sb.Append($"<tr><td style=\"padding-top:16px;color:#666;font-size:13px;\">{Encode(chainText)}</td></tr>");
        }

        // Buttons
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

        sb.Append("</table></td></tr></table></body></html>");

        return sb.ToString();
    }

    private static string Row(string label, string? value) =>
        $"<tr><td style=\"padding:4px 0;color:#666;width:140px;\">{Encode(label)}</td>" +
        $"<td style=\"padding:4px 0;\">{Encode(string.IsNullOrWhiteSpace(value) ? "-" : value)}</td></tr>";

    private static string Button(string url, string label, string colour) =>
        $"<a href=\"{Encode(url)}\" style=\"display:inline-block;padding:10px 20px;margin-right:8px;" +
        $"background:{colour};color:#ffffff;text-decoration:none;border-radius:4px;" +
        $"font-size:14px;font-weight:600;\">{Encode(label)}</a>";

    /// <summary>
    /// HTML-encodes every interpolated value. Vendor names come from a
    /// database somebody else can write to, and an unencoded apostrophe or
    /// angle bracket is how a broken layout becomes an injected link.
    /// </summary>
    private static string Encode(string? value) =>
        WebUtility.HtmlEncode(value ?? string.Empty);

    private static string FormatMoney(decimal amount, string? currencyCode)
    {
        var formatted = amount.ToString("N2", CultureInfo.InvariantCulture);
        return string.IsNullOrWhiteSpace(currencyCode) ? formatted : $"{currencyCode} {formatted}";
    }
}
