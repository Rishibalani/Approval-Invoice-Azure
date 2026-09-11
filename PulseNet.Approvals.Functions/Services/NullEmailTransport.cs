using Microsoft.Extensions.Logging;

namespace PulseNet.Approvals.Functions.Services;

/// <summary>
/// Logs the message instead of sending it. Used until a real transport is
/// chosen and configured.
///
/// Returns FAILURE rather than success. A transport that quietly pretends to
/// have delivered is worse than one that admits it cannot: the dispatcher
/// would mark the channel delivered, skip the fallback, and nobody would learn
/// that the approver was never told.
/// </summary>
public sealed class NullEmailTransport : IEmailTransport
{
    private readonly ILogger<NullEmailTransport> _logger;

    public string Name => "None";

    public NullEmailTransport(ILogger<NullEmailTransport> logger) => _logger = logger;

    public Task<EmailSendResult> SendAsync(EmailMessage message, CancellationToken cancellationToken)
    {
        _logger.LogWarning(
            "EMAIL NOT SENT - no transport configured.\nTo: {To}\nFrom: {From}\nSubject: {Subject}\n\n{Html}",
            message.ToAddress, message.FromAddress, message.Subject, message.HtmlBody);

        return Task.FromResult(EmailSendResult.Fail("email_transport_not_configured"));
    }
}