using System.Globalization;
using System.Net;
using System.Text;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Cards;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Options;
using PulseNet.Approvals.Functions.Security;
using PulseNet.Approvals.Functions.Services;

namespace PulseNet.Approvals.Functions.Channels;

/// <summary>
/// Sends an approval as an Outlook Actionable Message - an Adaptive Card
/// embedded in an email, with buttons that work inside the inbox.
///
/// HOW THE CARD GETS INTO THE EMAIL
///
/// Not as an attachment. The card JSON goes in a script block in the HTML
/// head:
///
///   &lt;script type="application/adaptivecard+json"&gt; { ... } &lt;/script&gt;
///
/// Outlook finds it, checks the originator against its registered providers,
/// and renders the card instead of the body. Anything that cannot render it -
/// mobile in some configurations, third-party clients, GCC High, DoD, personal
/// accounts - simply shows the HTML.
///
/// WHICH IS WHY THE HTML FALLBACK IS NOT OPTIONAL
///
/// It carries the same facts and a deep link into Business Central. An
/// approver whose client cannot render the card is then mildly inconvenienced
/// rather than completely stuck, and they never see an error - just an email
/// that works differently.
///
/// AUTHENTICATION CHANGED IN JUNE 2026
///
/// Legacy External Access Tokens for Actionable Messages were retired on
/// 8 June 2026. The action endpoint now validates a Microsoft Entra ID token.
/// Any sample describing a "Microsoft-issued bearer token" predates that and
/// no longer works - see ActionableMessageTokenValidator.
/// </summary>
public sealed class OutlookActionableMessageSender : IChannelSender
{
    private readonly IEmailTransport _transport;
    private readonly OutlookCardBuilder _cardBuilder;
    private readonly ActionTokenService _tokenService;
    private readonly ChannelOptions _options;
    private readonly ILogger<OutlookActionableMessageSender> _logger;

    public ChannelDeliveryMode Mode => ChannelDeliveryMode.ActionableMessage;
    public ApprovalChannel Channel => ApprovalChannel.Outlook;

    public OutlookActionableMessageSender(
        IEmailTransport transport,
        OutlookCardBuilder cardBuilder,
        ActionTokenService tokenService,
        IOptions<ChannelOptions> options,
        ILogger<OutlookActionableMessageSender> logger)
    {
        _transport = transport;
        _cardBuilder = cardBuilder;
        _tokenService = tokenService;
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
        var outlook = _options.Outlook;
        var recipient = payload.Approver.Upn;

        if (string.IsNullOrWhiteSpace(recipient))
        {
            // This is the fallback channel, so nothing downstream can rescue
            // it. Worth an error rather than a warning.
            _logger.LogError(
                "No UPN for approver {Approver}; the email channel cannot address anyone.",
                payload.Approver.UserId);

            return ChannelSendResult.Fail(Channel, "no_recipient_address");
        }

        // Without a registered originator Outlook silently declines to render
        // the card - no error, no card, just the HTML fallback. Degrading
        // deliberately is better than shipping a card nobody sees.
        if (string.IsNullOrWhiteSpace(outlook.OriginatorId))
        {
            _logger.LogWarning(
                "Channels:Outlook:OriginatorId is not set, so Outlook will not render the card. " +
                "Sending the HTML fallback with link buttons instead.");

            actionMode = actionMode == ChannelActionMode.Native
                ? ChannelActionMode.Link
                : actionMode;
        }

        try
        {
            // Tokens rather than URLs. Action.Http posts a body, so the token
            // travels in the payload instead of a query string - a body does
            // not land in browser history, proxy logs or referrer headers.
            string? approveToken = null;
            string? rejectToken = null;

            if (actionMode != ChannelActionMode.NotifyOnly &&
                !string.IsNullOrWhiteSpace(payload.Approver.Upn))
            {
                approveToken = _tokenService.Mint(
                    payload.Approval.ApprovalEntryNo, payload.Approver.Upn, ApprovalAction.Approve);

                rejectToken = _tokenService.Mint(
                    payload.Approval.ApprovalEntryNo, payload.Approver.Upn, ApprovalAction.Reject);
            }

            var html = BuildHtml(payload, actionMode, approveToken, rejectToken);

            var result = await _transport.SendAsync(new EmailMessage
            {
                ToAddress = recipient,
                ToDisplayName = payload.Approver.DisplayName,
                Subject = BuildSubject(payload),
                HtmlBody = html,
                FromAddress = outlook.FromAddress,
                FromDisplayName = outlook.FromDisplayName,
                OriginatorId = outlook.OriginatorId
            }, cancellationToken);

            if (!result.Succeeded)
            {
                return ChannelSendResult.Fail(
                    Channel, result.FailureReason ?? "send_failed", result.IsTransient);
            }

            _logger.LogInformation(
                "Outlook actionable message sent to {Recipient} for {DocumentNo}, action mode {ActionMode}.",
                recipient, payload.Document.DocumentNo, actionMode);

            return ChannelSendResult.Ok(Channel, result.MessageId);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Outlook sender threw for {DocumentNo}.", payload.Document.DocumentNo);
            return ChannelSendResult.Fail(Channel, "outlook_send_exception", transient: true);
        }
    }

    // ------------------------------------------------------------------

    private static string BuildSubject(ApprovalDispatchPayload payload)
    {
        var party = payload.Document.CounterpartyName ?? payload.Document.CounterpartyNo ?? "Unknown";

        var amount = FormatMoney(
            payload.Document.AmountInclTax != 0 ? payload.Document.AmountInclTax : payload.Document.Amount,
            payload.Document.CurrencyCode);

        return $"Approval needed: {payload.Document.DocumentNo} — {party} — {amount}";
    }

    /// <summary>
    /// The full HTML document: card in the head, fallback in the body.
    ///
    /// Table layout with inline styles throughout, which looks like 2005 web
    /// development because email clients are 2005 web browsers. Outlook renders
    /// HTML with Word's engine - no flexbox, no grid, no external stylesheets,
    /// unreliable div layout. Tables and inline styles are what survive.
    /// </summary>
    private string BuildHtml(
        ApprovalDispatchPayload payload,
        ChannelActionMode actionMode,
        string? approveToken,
        string? rejectToken)
    {
        var outlook = _options.Outlook;
        var sb = new StringBuilder();

        sb.Append("<html><head>");
        sb.Append("<meta charset=\"utf-8\">");

        // The card block. Outlook reads it; every other client ignores an
        // unknown script type and falls through to the body below.
        if (actionMode == ChannelActionMode.Native &&
            !string.IsNullOrWhiteSpace(outlook.OriginatorId))
        {
            var card = _cardBuilder.Build(
                payload,
                actionMode,
                outlook.OriginatorId,
                _options.ActionEndpointBaseUrl,
                approveToken,
                rejectToken);

            sb.Append("<script type=\"application/adaptivecard+json\">");
            sb.Append(card.ToJsonString());
            sb.Append("</script>");
        }

        sb.Append("</head>");
        sb.Append(BuildFallbackBody(payload, actionMode, approveToken, rejectToken));
        sb.Append("</html>");

        return sb.ToString();
    }

    private string BuildFallbackBody(
        ApprovalDispatchPayload payload,
        ChannelActionMode actionMode,
        string? approveToken,
        string? rejectToken)
    {
        var sb = new StringBuilder();
        var doc = payload.Document;
        var isPayable = doc.Direction == "Payable";
        var currency = doc.CurrencyCode;

        var headlineAmount = FormatMoney(
            doc.AmountInclTax != 0 ? doc.AmountInclTax : doc.Amount, currency);

        sb.Append("<body style=\"margin:0;padding:24px;background:#f5f5f5;");
        sb.Append("font-family:Segoe UI,Helvetica,Arial,sans-serif;color:#1a1a1a;\">");
        sb.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\"><tr><td align=\"center\">");
        sb.Append("<table role=\"presentation\" width=\"580\" cellpadding=\"0\" cellspacing=\"0\" ");
        sb.Append("style=\"background:#ffffff;border-radius:8px;padding:28px;\">");

        // Header
        sb.Append("<tr><td>");
        sb.Append($"<div style=\"font-size:18px;font-weight:600;\">{Enc(isPayable ? "Purchase invoice approval" : "Sales invoice approval")}</div>");
        sb.Append($"<div style=\"color:#666;margin-top:4px;\">{Enc(doc.CounterpartyName ?? doc.CounterpartyNo ?? "Unknown party")}</div>");
        sb.Append($"<div style=\"font-size:28px;font-weight:600;margin-top:16px;\">{Enc(headlineAmount)}</div>");
        sb.Append("</td></tr>");

        // Facts — same set and order as the card, so the two never disagree
        sb.Append("<tr><td style=\"padding-top:20px;\">");
        sb.Append("<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"font-size:14px;\">");

        sb.Append(Row("Document", doc.DocumentNo));
        sb.Append(Row("Their reference", doc.ExternalDocumentNo));
        sb.Append(Row(isPayable ? "Vendor" : "Customer", ComposeParty(doc.CounterpartyName, doc.CounterpartyNo)));

        if (doc.PayToDiffers)
        {
            sb.Append(Row("Pay-to", ComposeParty(doc.PayToName, doc.PayToNo)));
        }

        sb.Append(Row("Amount excl. tax", FormatMoney(doc.AmountExclTax, currency)));

        if (doc.AmountInclTax != doc.AmountExclTax)
        {
            sb.Append(Row("Amount incl. tax", FormatMoney(doc.AmountInclTax, currency)));
        }

        if (!string.IsNullOrWhiteSpace(currency) && doc.Amount != doc.AmountLcy)
        {
            sb.Append(Row("Local value", FormatMoney(doc.AmountLcy, null)));
        }

        sb.Append(Row("Document date", doc.DocumentDate));
        sb.Append(Row("Posting date", doc.PostingDate));
        sb.Append(Row("Due", doc.DueDate));
        sb.Append(Row("Requested by", payload.Approval.RequesterLabel));

        sb.Append("</table></td></tr>");

        // Warnings, above the buttons — an approver who has already decided by
        // the time they scroll past will not read a caveat placed underneath.
        if (doc.PayToDiffers)
        {
            sb.Append(Warning("Payment goes to a different party than the vendor on this invoice. Worth confirming that is expected.", "#d97706"));
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
                sb.Append(Warning(text, "#a4262c"));
            }
        }

        // Chain context
        if (payload.Approval.TotalStepsInChain > 1)
        {
            var chain = payload.Approval.IsFinalStep
                ? $"Approval {payload.Approval.SequenceNo} of {payload.Approval.TotalStepsInChain}. This is the final approval — approving releases the invoice."
                : $"Approval {payload.Approval.SequenceNo} of {payload.Approval.TotalStepsInChain}. Further approval is required after yours.";

            sb.Append($"<tr><td style=\"padding-top:16px;color:#666;font-size:13px;\">{Enc(chain)}</td></tr>");
        }

        // Buttons. In the fallback these are always LINKS, never Action.Http -
        // a plain client cannot POST. They carry the same token in a query
        // string and land on the browser endpoint instead.
        sb.Append("<tr><td style=\"padding-top:24px;\">");

        if (actionMode != ChannelActionMode.NotifyOnly)
        {
            if (!string.IsNullOrWhiteSpace(approveToken))
            {
                sb.Append(Button(LinkFor(approveToken), "Approve", "#107c10"));
            }

            if (!string.IsNullOrWhiteSpace(rejectToken))
            {
                sb.Append(Button(LinkFor(rejectToken), "Reject", "#a4262c"));
            }
        }

        if (!string.IsNullOrWhiteSpace(doc.DeepLink))
        {
            sb.Append(Button(doc.DeepLink, "View in Business Central", "#5a5a5a"));
        }

        sb.Append("</td></tr>");

        sb.Append("<tr><td style=\"padding-top:24px;color:#999;font-size:12px;\">");
        sb.Append($"Approval buttons expire {payload.Policy.ActionTokenTtlMinutes} minutes after this was sent. ");
        sb.Append("After that, please use Business Central.");
        sb.Append("</td></tr>");

        sb.Append("</table></td></tr></table></body>");

        return sb.ToString();
    }

    private string LinkFor(string token)
    {
        var baseUrl = _options.ActionEndpointBaseUrl;
        var separator = baseUrl.Contains('?') ? "&" : "?";
        return $"{baseUrl}{separator}t={Uri.EscapeDataString(token)}";
    }

    private static string Row(string label, string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return $"<tr><td style=\"padding:4px 0;color:#666;width:150px;vertical-align:top;\">{Enc(label)}</td>" +
               $"<td style=\"padding:4px 0;\">{Enc(value)}</td></tr>";
    }

    private static string Warning(string text, string colour) =>
        "<tr><td style=\"padding-top:16px;\">" +
        $"<div style=\"background:#fff4e5;border-left:3px solid {colour};padding:12px;font-size:14px;\">" +
        Enc(text) + "</div></td></tr>";

    private static string Button(string url, string label, string colour) =>
        $"<a href=\"{Enc(url)}\" style=\"display:inline-block;padding:10px 20px;margin-right:8px;" +
        $"background:{colour};color:#ffffff;text-decoration:none;border-radius:4px;" +
        $"font-size:14px;font-weight:600;\">{Enc(label)}</a>";

    /// <summary>
    /// HTML-encodes every interpolated value. Vendor names come from a database
    /// somebody else can write to, and an unencoded angle bracket is how a
    /// broken layout becomes an injected link.
    /// </summary>
    private static string Enc(string? value) => WebUtility.HtmlEncode(value ?? string.Empty);

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
