using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Options;
using PulseNet.Approvals.Functions.Security;

namespace PulseNet.Approvals.Functions.Channels;

/// <summary>
/// Fans one payload out across the channels Business Central asked for.
///
/// WHO DECIDES WHAT
///
/// Business Central decides WHICH channels. The global toggles on the
/// Approval Channel Setup page produce payload.Approver.Channels, and this
/// class sends to exactly those and nothing else. It has no opinion of its own
/// about which channels are in use, and no configuration here can add one
/// Business Central did not ask for.
///
/// Azure decides HOW each channel is reached - WorkflowWebhook today, Bot
/// later. That is transport, not policy.
///
/// The split matters because it keeps one switch per decision. An admin who
/// turns Teams off in Business Central does not then have to find someone with
/// Azure access to make it stick.
///
/// When Business Central names a channel Azure has no transport for, that is
/// logged as an error naming both sides rather than skipped quietly - a
/// channel that is on in Business Central and invisible in Azure is exactly
/// the failure that wastes an afternoon.
///
/// TWO RULES OTHERWISE
///
/// 1. Every channel is isolated. One failing never stops another being tried.
/// 2. Policy comes from Business Central. This class reads the verdict and
///    downgrades to NotifyOnly when told to. It never upgrades, and it never
///    works out for itself whether an amount is too large.
/// </summary>
public sealed class ChannelDispatcher
{
    private readonly IEnumerable<IChannelSender> _senders;
    private readonly ActionTokenService _tokenService;
    private readonly ChannelOptions _options;
    private readonly ILogger<ChannelDispatcher> _logger;

    public ChannelDispatcher(
        IEnumerable<IChannelSender> senders,
        ActionTokenService tokenService,
        IOptions<ChannelOptions> options,
        ILogger<ChannelDispatcher> logger)
    {
        _senders = senders;
        _tokenService = tokenService;
        _options = options.Value;
        _logger = logger;
    }

    public async Task<DispatchOutcome> DispatchAsync(
        ApprovalDispatchPayload payload,
        CancellationToken cancellationToken)
    {
        var results = new List<ChannelSendResult>();

        // Suspended means on leave or opted out entirely. Their approvals still
        // work in Business Central; they just do not get chased in a channel.
        if (payload.Approver.Suspended)
        {
            _logger.LogInformation(
                "Approver {Approver} is suspended. Nothing dispatched.",
                payload.Approver.UserId);

            return new DispatchOutcome
            {
                Results = [ChannelSendResult.Skipped(ApprovalChannel.None, "approver_suspended")],
                AnySucceeded = false,
                WasSuppressed = true
            };
        }

        // Business Central's list. Empty means every channel is switched off
        // there, which is a deliberate state and not an error.
        var requested = ResolveRequestedChannels(payload);

        if (requested.Count == 0)
        {
            _logger.LogInformation(
                "Business Central has no channels enabled. {DocumentNo} not dispatched.",
                payload.Document.DocumentNo);

            return new DispatchOutcome
            {
                Results = [ChannelSendResult.Skipped(ApprovalChannel.None, "no_channels_enabled_in_bc")],
                AnySucceeded = false,
                WasSuppressed = true
            };
        }

        foreach (var channel in requested)
        {
            results.Add(await TrySendAsync(payload, channel, cancellationToken));
        }

        // Fallback ladder. Only entered when every requested channel failed.
        if (!results.Any(r => r.Succeeded))
        {
            var fallback = await TryFallbackAsync(payload, results, cancellationToken);

            if (fallback is not null)
            {
                results.Add(fallback);
            }
        }

        var anySucceeded = results.Any(r => r.Succeeded);

        if (!anySucceeded)
        {
            // Nobody will approve this from a channel. Level 1 escalation - the
            // only failure an approver genuinely needs a human to chase.
            _logger.LogError(
                "Every channel failed for {DocumentNo}, approver {Approver}. Reasons: {Reasons}",
                payload.Document.DocumentNo,
                payload.Approver.UserId,
                string.Join(", ", results.Select(r => $"{r.Channel}={r.FailureReason}")));
        }

        return new DispatchOutcome { Results = results, AnySucceeded = anySucceeded };
    }

    // ------------------------------------------------------------------

    /// <summary>
    /// Turns Business Central's channel names into enum values. Unknown names
    /// are logged and ignored rather than throwing - a future Business Central
    /// build adding a channel this Function does not know about should degrade,
    /// not break.
    /// </summary>
    private List<ApprovalChannel> ResolveRequestedChannels(ApprovalDispatchPayload payload)
    {
        var resolved = new List<ApprovalChannel>();

        foreach (var name in payload.Approver.Channels)
        {
            if (Enum.TryParse<ApprovalChannel>(name, ignoreCase: true, out var channel) &&
                channel != ApprovalChannel.None)
            {
                resolved.Add(channel);
            }
            else
            {
                _logger.LogWarning(
                    "Business Central asked for channel '{Channel}', which this build does not implement. Ignoring.",
                    name);
            }
        }

        return resolved;
    }

    private async Task<ChannelSendResult> TrySendAsync(
        ApprovalDispatchPayload payload,
        ApprovalChannel channel,
        CancellationToken cancellationToken)
    {
        var (delivery, configuredActionMode) = GetTransport(channel);

        // Business Central asked for this channel but Azure has no transport
        // configured. Loud, because the symptom otherwise is a notification
        // that simply never arrives with nothing anywhere explaining why.
        if (delivery == ChannelDeliveryMode.Disabled)
        {
            _logger.LogError(
                "Business Central has {Channel} switched on, but Channels:{Channel}:DeliveryMode is Disabled in this Function. " +
                "Set a delivery mode, or switch {Channel} off on the Approval Channel Setup page.",
                channel, channel, channel);

            return ChannelSendResult.Fail(channel, "transport_not_configured");
        }

        var sender = _senders.FirstOrDefault(s => s.Mode == delivery);

        if (sender is null)
        {
            _logger.LogError(
                "{Channel} is configured for {Mode} but no sender implements it.",
                channel, delivery);

            return ChannelSendResult.Fail(channel, $"no_sender_for_{delivery}");
        }

        // ---- Policy ---------------------------------------------------
        // The only place the verdict is applied. Downgrade only.
        var policy = payload.Policy.ResolveFor(channel.ToString());
        var actionMode = policy.CanApprove ? configuredActionMode : ChannelActionMode.NotifyOnly;

        if (!policy.CanApprove)
        {
            _logger.LogInformation(
                "{Channel} downgraded to notify-only for {DocumentNo}: {Reasons}",
                channel, payload.Document.DocumentNo, string.Join(",", policy.Reasons));
        }

        // ---- Tokens ---------------------------------------------------
        string? approveUrl = null;
        string? rejectUrl = null;

        if (actionMode == ChannelActionMode.Link)
        {
            if (string.IsNullOrWhiteSpace(payload.Approver.Upn))
            {
                // The token binds to a hashed UPN. Without one there is nothing
                // to bind to, so a link would be an unbound approval button -
                // strictly worse than no button at all.
                _logger.LogWarning(
                    "No UPN for {Approver}; falling back to notify-only on {Channel}.",
                    payload.Approver.UserId, channel);

                actionMode = ChannelActionMode.NotifyOnly;
            }
            else if (string.IsNullOrWhiteSpace(_options.ActionEndpointBaseUrl))
            {
                _logger.LogError("Channels:ActionEndpointBaseUrl is not set. Buttons suppressed.");
                actionMode = ChannelActionMode.NotifyOnly;
            }
            else
            {
                // Separate tokens, separate nonces. Burning Approve must not
                // silently disable Reject.
                approveUrl = _tokenService.BuildActionUrl(
                    _options.ActionEndpointBaseUrl,
                    _tokenService.Mint(payload.Approval.ApprovalEntryNo, payload.Approver.Upn, ApprovalAction.Approve));

                rejectUrl = _tokenService.BuildActionUrl(
                    _options.ActionEndpointBaseUrl,
                    _tokenService.Mint(payload.Approval.ApprovalEntryNo, payload.Approver.Upn, ApprovalAction.Reject));
            }
        }

        // ---- Send -----------------------------------------------------
        // Isolated. A sender that throws despite the contract must not take the
        // other channels down with it.
        try
        {
            return await sender.SendAsync(payload, actionMode, approveUrl, rejectUrl, cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "{Channel} sender threw unexpectedly.", channel);
            return ChannelSendResult.Fail(channel, "sender_exception", transient: true);
        }
    }

    private async Task<ChannelSendResult?> TryFallbackAsync(
        ApprovalDispatchPayload payload,
        List<ChannelSendResult> alreadyTried,
        CancellationToken cancellationToken)
    {
        // Business Central names the fallback too. Azure only supplies the
        // transport for it - and, when the payload omits the field entirely,
        // the configured Channels:FallbackChannel.
        ApprovalChannel fallback;

        if (payload.Approver.FallbackChannel is null)
        {
            fallback = _options.FallbackChannel;
        }
        else if (!Enum.TryParse<ApprovalChannel>(payload.Approver.FallbackChannel, ignoreCase: true, out fallback))
        {
            return null;
        }

        if (fallback == ApprovalChannel.None)
        {
            return null;
        }

        if (alreadyTried.Any(r => r.Channel == fallback))
        {
            // Already attempted and failed. Retrying the same transport
            // immediately would just fail the same way.
            return null;
        }

        _logger.LogWarning(
            "All requested channels failed for {DocumentNo}. Trying fallback {Fallback}.",
            payload.Document.DocumentNo, fallback);

        return await TrySendAsync(payload, fallback, cancellationToken);
    }

    /// <summary>
    /// How this channel is physically reached. Transport only - this never
    /// decides whether a channel is used, only how.
    /// </summary>
    private (ChannelDeliveryMode Delivery, ChannelActionMode Action) GetTransport(ApprovalChannel channel) =>
        channel switch
        {
            ApprovalChannel.Teams => (_options.Teams.DeliveryMode, _options.Teams.ActionMode),
            ApprovalChannel.Outlook => (_options.Outlook.DeliveryMode, _options.Outlook.ActionMode),
            ApprovalChannel.WhatsApp => (_options.WhatsApp.DeliveryMode, _options.WhatsApp.ActionMode),
            _ => (ChannelDeliveryMode.Disabled, ChannelActionMode.NotifyOnly)
        };
}

public sealed record DispatchOutcome
{
    public required IReadOnlyList<ChannelSendResult> Results { get; init; }
    public required bool AnySucceeded { get; init; }
    public bool WasSuppressed { get; init; }

    public string Summary =>
        string.Join(", ", Results.Select(r => $"{r.Channel}={(r.Succeeded ? "ok" : r.FailureReason)}"));
}
