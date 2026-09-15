using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Cards;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Options;
using PulseNet.Approvals.Functions.Services;

namespace PulseNet.Approvals.Functions.Channels;

/// <summary>
/// Plain HTML email with link buttons. The universal fallback.
///
/// Every approver has a mailbox. Not every approver has the Teams app
/// installed, a verified WhatsApp number, or a working Actionable Messages
/// registration. This is the channel that always resolves to something.
///
/// WHAT IT SHARES, AND WHY
///
/// The body comes from ApprovalEmailBuilder, which reads
/// ApprovalCardViewModel - the same source as the Teams card and the Outlook
/// card. An earlier version formatted its own money and rendered a missing
/// requester as a bare dash, so the same invoice could read three different
/// ways depending on which delivery mode happened to be configured.
///
/// The transport comes from IEmailTransport, chosen by configuration. Which
/// one to use is a decision with cost and permission consequences, and it does
/// not belong entangled with what the email says.
/// </summary>
public sealed class PlainEmailSender : IChannelSender
{
    private readonly IEmailTransport _transport;
    private readonly ApprovalEmailBuilder _emailBuilder;
    private readonly ChannelOptions _options;
    private readonly ILogger<PlainEmailSender> _logger;

    public ChannelDeliveryMode Mode => ChannelDeliveryMode.PlainEmail;
    public ApprovalChannel Channel => ApprovalChannel.Outlook;

    public PlainEmailSender(
        IEmailTransport transport,
        ApprovalEmailBuilder emailBuilder,
        IOptions<ChannelOptions> options,
        ILogger<PlainEmailSender> logger)
    {
        _transport = transport;
        _emailBuilder = emailBuilder;
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
            // This is the fallback, so nothing downstream can rescue it.
            // Worth an error rather than a warning.
            _logger.LogError(
                "No UPN for approver {Approver}; the email fallback cannot deliver.",
                payload.Approver.UserId);

            return ChannelSendResult.Fail(Channel, "no_recipient_address");
        }

        var subject = _emailBuilder.BuildSubject(payload);

        var html = "<html><head><meta charset=\"utf-8\"></head>"
                   + _emailBuilder.BuildBody(payload, actionMode, approveUrl, rejectUrl)
                   + "</html>";

        var result = await _transport.SendAsync(new EmailMessage
        {
            ToAddress = recipient,
            ToDisplayName = payload.Approver.DisplayName,
            Subject = subject,
            HtmlBody = html,
            FromAddress = _options.Outlook.FromAddress,
            FromDisplayName = _options.Outlook.FromDisplayName
        }, cancellationToken);

        if (!result.Succeeded)
        {
            return ChannelSendResult.Fail(
                Channel, result.FailureReason ?? "send_failed", result.IsTransient);
        }

        _logger.LogInformation(
            "Plain approval email sent to {Recipient} for {DocumentNo}.",
            recipient, payload.Document.DocumentNo);

        return ChannelSendResult.Ok(Channel, result.MessageId);
    }
}
