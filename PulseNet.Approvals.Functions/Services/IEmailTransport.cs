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
namespace PulseNet.Approvals.Functions.Services;

/// <summary>
/// Puts a message on the wire. Deliberately separate from the senders, because
/// which transport to use is a real decision with cost and permission
/// consequences, and it should not be entangled with what the email says.
///
///   Graph   Mail.Send application permission plus admin consent, and a
///           mailbox to send as. Mail lands in that mailbox's Sent Items,
///           which auditors like. Narrow it with an Exchange application
///           access policy - the raw permission allows sending as ANY mailbox
///           in the tenant, which is not a thing to leave unrestricted.
///
///   ACS     No Graph permission, no admin consent. Costs per message, and
///           sends from a service domain unless you verify your own.
///
///   Null    Logs the message. Used until one of the above is chosen, and
///           deliberately loud rather than silently pretending to succeed.
/// </summary>
public interface IEmailTransport
{
    string Name { get; }

    Task<EmailSendResult> SendAsync(EmailMessage message, CancellationToken cancellationToken);
}

public sealed record EmailMessage
{
    public required string ToAddress { get; init; }
    public string? ToDisplayName { get; init; }
    public required string Subject { get; init; }
    public required string HtmlBody { get; init; }
    public required string FromAddress { get; init; }
    public string? FromDisplayName { get; init; }

    /// <summary>
    /// Set on Actionable Messages. Some Outlook configurations use it as an
    /// additional signal that the sender is a registered provider.
    /// </summary>
    public string? OriginatorId { get; init; }
}

public sealed record EmailSendResult
{
    public required bool Succeeded { get; init; }
    public string? MessageId { get; init; }
    public string? FailureReason { get; init; }
    public bool IsTransient { get; init; }

    public static EmailSendResult Ok(string? messageId = null) =>
        new() { Succeeded = true, MessageId = messageId };

    public static EmailSendResult Fail(string reason, bool transient = false) =>
        new() { Succeeded = false, FailureReason = reason, IsTransient = transient };
}

*/
