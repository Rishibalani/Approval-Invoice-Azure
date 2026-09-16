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
﻿using Microsoft.Extensions.Logging;

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
*/
