using System.Text.Json;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using PulseNet.Approvals.Functions.Channels;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Services;

namespace PulseNet.Approvals.Functions.Functions;

/// <summary>
/// Reads a queued payload and fans it out to the channels.
///
/// WHY THIS IS BEHIND A QUEUE RATHER THAN INSIDE THE INGEST FUNCTION
///
/// A channel outage must never become a Business Central retry storm. Ingest
/// returns 202 the moment the payload is durably parked; Business Central
/// marks the outbox row Sent and stops caring. If Teams is down, the retry
/// happens here, on the queue's own schedule, close to the actual failure.
///
/// The retry loop that matters is always the one nearest the thing that broke.
///
/// POISON HANDLING IS FREE
///
/// After maxDequeueCount attempts (the host default of 5, overridable under
/// extensions.queues in host.json) the message lands on the queue named
/// {Dispatch:QueueName}-poison. Put an alert on that queue's length - it is
/// the level 4 escalation from the plan.
/// </summary>
public sealed class ApprovalDispatchWorker
{
    private readonly ChannelDispatcher _dispatcher;
    private readonly BusinessCentralClient _bcClient;
    private readonly IdempotencyStore _idempotencyStore;
    private readonly CardRefreshService _cardRefresh;
    private readonly ILogger<ApprovalDispatchWorker> _logger;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public ApprovalDispatchWorker(
        ChannelDispatcher dispatcher,
        BusinessCentralClient bcClient,
        IdempotencyStore idempotencyStore,
        CardRefreshService cardRefresh,
        ILogger<ApprovalDispatchWorker> logger)
    {
        _dispatcher = dispatcher;
        _bcClient = bcClient;
        _idempotencyStore = idempotencyStore;
        _cardRefresh = cardRefresh;
        _logger = logger;
    }

    [Function(nameof(ApprovalDispatchWorker))]
    public async Task Run(
        [QueueTrigger("%Dispatch:QueueName%", Connection = "AzureWebJobsStorage")]
        string message,
        CancellationToken cancellationToken)
    {
        ApprovalDispatchPayload? payload;

        try
        {
            payload = JsonSerializer.Deserialize<ApprovalDispatchPayload>(message, JsonOptions);
        }
        catch (JsonException ex)
        {
            // Throwing sends it to the poison queue, which is what we want -
            // a message we cannot parse will never parse, and retrying it
            // five times just burns money.
            _logger.LogError(ex, "Queue message could not be deserialised.");
            throw;
        }

        if (payload is null)
        {
            throw new InvalidOperationException("Queue message deserialised to null.");
        }

        using var scope = _logger.BeginScope(new Dictionary<string, object>
        {
            ["CorrelationId"] = payload.CorrelationId,
            ["ApprovalEntryNo"] = payload.Approval.ApprovalEntryNo,
            ["DocumentNo"] = payload.Document.DocumentNo
        });

        // Status-change events retire or update an existing card rather than
        // raising a new one. Until in-place update exists (bot mode), logging
        // them is honest: we know, and there is nothing yet to update.
        if (payload.EventType != "Requested")
        {
            // A decision was made somewhere - here, in another channel, or in
            // Business Central. Replace the card so it stops looking like it
            // is waiting for one.
            //
            // Teams only. An email that has left cannot be changed, and a
            // second email would be worse than a stale one.
            _logger.LogInformation(
                "{EventType} for {DocumentNo}. Refreshing any card already sent.",
                payload.EventType, payload.Document.DocumentNo);

            await _cardRefresh.RefreshAsync(payload, cancellationToken);

            await _idempotencyStore.MarkCompletedAsync(
                payload.EventId, payload.Source.TenantId, "CardRefreshed", cancellationToken);

            return;
        }

        var outcome = await _dispatcher.DispatchAsync(payload, cancellationToken);

        _logger.LogInformation(
            "Dispatched {DocumentNo} to {Approver}: {Summary}",
            payload.Document.DocumentNo, payload.Approver.UserId, outcome.Summary);

        // Correlate the delivery so the card can be updated in place later.
        foreach (var success in outcome.Results.Where(r => r.Succeeded))
        {
            await _bcClient.RecordDeliveryAsync(
                payload.EventId, success.Channel, success.ChannelMessageId, cancellationToken);
        }

        await _idempotencyStore.MarkCompletedAsync(
            payload.EventId,
            payload.Source.TenantId,
            outcome.AnySucceeded ? "Delivered" : "AllChannelsFailed",
            cancellationToken);

        // Throwing here would retry the WHOLE fan-out, re-sending to channels
        // that already succeeded. Only retry when nothing got through and the
        // failure looks transient.
        if (!outcome.AnySucceeded && !outcome.WasSuppressed)
        {
            var anyTransient = outcome.Results.Any(r => r.IsTransient);

            if (anyTransient)
            {
                throw new InvalidOperationException(
                    $"All channels failed transiently for {payload.Document.DocumentNo}. Requeuing.");
            }

            // Permanent failure - misconfiguration, no recipient address.
            // Retrying will not help; leave it for a human.
            _logger.LogError(
                "All channels failed permanently for {DocumentNo}: {Summary}",
                payload.Document.DocumentNo, outcome.Summary);
        }
    }
}
