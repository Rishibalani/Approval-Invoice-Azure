namespace PulseNet.Approvals.Functions.Models;

/// <summary>
/// Outcome of one channel's send attempt.
///
/// Deliberately not an exception type. A channel failing is expected and
/// routine - the dispatcher tries the next channel and only escalates when
/// every one of them has failed. Exceptions are for things that should not
/// happen; a Teams outage is not one of those.
/// </summary>
public sealed record ChannelSendResult
{
    public required ApprovalChannel Channel { get; init; }
    public required bool Succeeded { get; init; }

    /// <summary>
    /// Identifier the channel assigned to the message. Needed later to update
    /// the card in place once the approval is decided. Null when the channel
    /// does not return one (a Workflows webhook does not).
    /// </summary>
    public string? ChannelMessageId { get; init; }

    public string? FailureReason { get; init; }

    /// <summary>True when retrying might work: a timeout, a 5xx, a throttle.</summary>
    public bool IsTransient { get; init; }

    public static ChannelSendResult Ok(ApprovalChannel channel, string? messageId = null) =>
        new() { Channel = channel, Succeeded = true, ChannelMessageId = messageId };

    public static ChannelSendResult Fail(ApprovalChannel channel, string reason, bool transient = false) =>
        new() { Channel = channel, Succeeded = false, FailureReason = reason, IsTransient = transient };

    public static ChannelSendResult Skipped(ApprovalChannel channel, string reason) =>
        new() { Channel = channel, Succeeded = false, FailureReason = "skipped: " + reason };
}
