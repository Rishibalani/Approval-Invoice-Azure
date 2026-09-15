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
    private readonly ApprovalEmailBuilder _emailBuilder;
    private readonly ActionTokenService _tokenService;
    private readonly ChannelOptions _options;
    private readonly ILogger<OutlookActionableMessageSender> _logger;

    public ChannelDeliveryMode Mode => ChannelDeliveryMode.ActionableMessage;
    public ApprovalChannel Channel => ApprovalChannel.Outlook;

    public OutlookActionableMessageSender(
        IEmailTransport transport,
        OutlookCardBuilder cardBuilder,
        ApprovalEmailBuilder emailBuilder,
        ActionTokenService tokenService,
        IOptions<ChannelOptions> options,
        ILogger<OutlookActionableMessageSender> logger)
    {
        _transport = transport;
        _cardBuilder = cardBuilder;
        _emailBuilder = emailBuilder;
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
                Subject = _emailBuilder.BuildSubject(payload),
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

    /// <summary>
    /// The full HTML document: card in the head, fallback in the body.
    ///
    /// Outlook finds the script block, checks the originator against its
    /// registered providers, and renders the card instead of the body.
    /// Anything that cannot - GCC High, DoD, personal accounts, third-party
    /// clients, some mobile configurations - simply shows the HTML.
    ///
    /// Which is why the fallback is not optional. It carries the same facts,
    /// from the same view model, so the two can never disagree about what an
    /// invoice says.
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

        // In the fallback the buttons are always LINKS, never Action.Http - a
        // plain client cannot POST. Same token, query string instead of body,
        // landing on the browser endpoint.
        sb.Append(_emailBuilder.BuildBody(
            payload,
            actionMode,
            approveToken is null ? null : LinkFor(approveToken),
            rejectToken is null ? null : LinkFor(rejectToken)));

        sb.Append("</html>");

        return sb.ToString();
    }

    private string LinkFor(string token)
    {
        var baseUrl = _options.ActionEndpointBaseUrl;
        var separator = baseUrl.Contains('?') ? "&" : "?";
        return $"{baseUrl}{separator}t={Uri.EscapeDataString(token)}";
    }
}
