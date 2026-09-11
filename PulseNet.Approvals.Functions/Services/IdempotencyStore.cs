using Azure;
using Azure.Data.Tables;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Options;

namespace PulseNet.Approvals.Functions.Services;

/// <summary>
/// Deduplication on the idempotency key Business Central sends with every
/// attempt, including retries.
///
/// This is not optional. A timeout is indistinguishable from a slow success:
/// Business Central cannot tell whether the Function processed the payload
/// before the connection dropped, so it retries. Without this store, one
/// network blip means the approver gets two identical cards for the same
/// invoice - and if they tap both, two people spend an afternoon working out
/// whether the invoice was approved once or twice.
///
/// The insert IS the check. TryClaim relies on Table Storage rejecting a
/// duplicate row key with a 409, so two concurrent retries cannot both win.
/// Compare-then-insert would have a race; this does not.
/// </summary>
public sealed class IdempotencyStore
{
    private readonly TableServiceClient _tableService;
    private readonly DispatchOptions _options;
    private readonly ILogger<IdempotencyStore> _logger;

    public IdempotencyStore(
        TableServiceClient tableService,
        IOptions<DispatchOptions> options,
        ILogger<IdempotencyStore> logger)
    {
        _tableService = tableService;
        _options = options.Value;
        _logger = logger;
    }

    /// <summary>
    /// Attempts to claim an event. Returns true if this is the first sighting
    /// and processing should continue; false if it is a duplicate.
    /// </summary>
    public async Task<bool> TryClaimAsync(
        string eventId,
        string tenantId,
        string correlationId,
        CancellationToken cancellationToken)
    {
        var table = _tableService.GetTableClient(_options.IdempotencyTable);
        await table.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

        // Partitioning on tenant keeps one noisy tenant from hot-spotting
        // another's partition, and makes a per-tenant purge trivial.
        var entity = new TableEntity(SanitisePartition(tenantId), eventId)
        {
            ["CorrelationId"] = correlationId,
            ["ClaimedUtc"] = DateTime.UtcNow,
            ["Status"] = "Claimed"
        };

        try
        {
            await table.AddEntityAsync(entity, cancellationToken);
            return true;
        }
        catch (RequestFailedException ex) when (ex.Status == 409)
        {
            _logger.LogInformation(
                "Duplicate dispatch discarded. EventId={EventId} CorrelationId={CorrelationId}",
                eventId, correlationId);
            return false;
        }
    }

    /// <summary>
    /// Marks a claimed event as finished. Purely for diagnostics - a row stuck
    /// on Claimed means the worker died between claim and completion, which is
    /// worth a query when something looks wrong.
    /// </summary>
    public async Task MarkCompletedAsync(
        string eventId,
        string tenantId,
        string outcome,
        CancellationToken cancellationToken)
    {
        var table = _tableService.GetTableClient(_options.IdempotencyTable);

        try
        {
            var entity = new TableEntity(SanitisePartition(tenantId), eventId)
            {
                ["ClaimedUtc"] = DateTime.UtcNow,
                ["Status"] = outcome,
                ["CompletedUtc"] = DateTime.UtcNow
            };

            await table.UpdateEntityAsync(entity, ETag.All, TableUpdateMode.Merge, cancellationToken);
        }
        catch (RequestFailedException ex)
        {
            // Never let a bookkeeping failure fail the actual work.
            _logger.LogWarning(ex, "Could not update idempotency row for {EventId}.", eventId);
        }
    }

    // Table Storage rejects / \ # ? and control characters in keys.
    private static string SanitisePartition(string value) =>
        string.IsNullOrWhiteSpace(value)
            ? "unknown"
            : value.Replace('/', '-').Replace('\\', '-').Replace('#', '-').Replace('?', '-');
}
