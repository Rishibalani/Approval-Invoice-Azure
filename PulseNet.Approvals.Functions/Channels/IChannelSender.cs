using PulseNet.Approvals.Functions.Models;

namespace PulseNet.Approvals.Functions.Channels;

/// <summary>
/// One implementation per delivery mode.
///
/// The whole point of this interface is that adding WhatsApp, or swapping the
/// Teams webhook for a real bot, is a new class and a config change - never a
/// change to the dispatch pipeline. ChannelDispatcher does not know how many
/// implementations exist, and does not care.
/// </summary>
public interface IChannelSender
{
    /// <summary>Which delivery mode this sender implements.</summary>
    ChannelDeliveryMode Mode { get; }

    /// <summary>Which logical channel it delivers to, for logging and audit.</summary>
    ApprovalChannel Channel { get; }

    /// <summary>
    /// Sends one approval notification.
    ///
    /// MUST NOT THROW for a delivery failure. Return a failed
    /// ChannelSendResult instead - the dispatcher then tries the next channel
    /// and only escalates when every one has failed. A Teams outage is an
    /// expected condition, not an exceptional one.
    /// </summary>
    Task<ChannelSendResult> SendAsync(
        ApprovalDispatchPayload payload,
        ChannelActionMode actionMode,
        string? approveUrl,
        string? rejectUrl,
        CancellationToken cancellationToken);
}
