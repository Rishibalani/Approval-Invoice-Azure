using Azure;
using Azure.Data.Tables;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Options;

namespace PulseNet.Approvals.Functions.Services;

/// <summary>
/// Remembers which Teams message carried which approval, so the card can be
/// replaced once a decision is made anywhere.
///
/// WHY THIS LIVES IN AZURE RATHER THAN BUSINESS CENTRAL
///
/// The outbox table has a Channel Message ID field, and writing back to it
/// would be the tidier home. But it needs an API page, a permission, and a
/// round trip on every send - and the value is only ever read by the code that
/// wrote it, moments later, in the same service.
///
/// Table Storage costs nothing, needs no schema change in Business Central,
/// and keeps a purely mechanical detail out of the business data.
///
/// KEYED ON APPROVAL ENTRY NUMBER
///
/// That is what a status event carries. The outbox event id would be more
/// precise, but a Rejected event has a different event id from the Requested
/// one that sent the card - so keying on it would never match.
///
/// One consequence: if a card is somehow sent twice for the same entry, the
/// second overwrites the first and only the newer one is refreshed. Acceptable
/// - duplicate sends are already prevented by the idempotency store, and the
/// stale card's buttons still fail safely.
/// </summary>
public sealed class SentCardStore
{
    private const string Partition = "card";

    private readonly TableServiceClient _tableService;
    private readonly TeamsBotOptions _options;
    private readonly ILogger<SentCardStore> _logger;

    public SentCardStore(
        TableServiceClient tableService,
        IOptions<TeamsBotOptions> options,
        ILogger<SentCardStore> logger)
    {
        _tableService = tableService;
        _options = options.Value;
        _logger = logger;
    }

    private string TableName => _options.SentCardTable;

    public async Task SaveAsync(SentCard card, CancellationToken cancellationToken)
    {
        try
        {
            var table = _tableService.GetTableClient(TableName);
            await table.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

            var entity = new TableEntity(Partition, card.ApprovalEntryNo.ToString())
            {
                ["ConversationId"] = card.ConversationId,
                ["ActivityId"] = card.ActivityId,
                ["ServiceUrl"] = card.ServiceUrl,
                ["DocumentNo"] = card.DocumentNo,
                ["SentUtc"] = DateTime.UtcNow
            };

            await table.UpsertEntityAsync(entity, TableUpdateMode.Replace, cancellationToken);
        }
        catch (Exception ex)
        {
            // Never fail a send because the bookkeeping failed. The approver
            // still has a working card; it simply will not refresh itself.
            _logger.LogWarning(ex, "Could not record the sent card for entry {EntryNo}.", card.ApprovalEntryNo);
        }
    }

    public async Task<SentCard?> GetAsync(int approvalEntryNo, CancellationToken cancellationToken)
    {
        try
        {
            var table = _tableService.GetTableClient(TableName);
            await table.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

            var entity = await table.GetEntityAsync<TableEntity>(
                Partition, approvalEntryNo.ToString(), cancellationToken: cancellationToken);

            return new SentCard
            {
                ApprovalEntryNo = approvalEntryNo,
                ConversationId = entity.Value.GetString("ConversationId") ?? string.Empty,
                ActivityId = entity.Value.GetString("ActivityId") ?? string.Empty,
                ServiceUrl = entity.Value.GetString("ServiceUrl") ?? string.Empty,
                DocumentNo = entity.Value.GetString("DocumentNo") ?? string.Empty
            };
        }
        catch (RequestFailedException ex) when (ex.Status == 404)
        {
            // No card was sent for this entry, or it went out before this
            // store existed. Nothing to refresh, and nothing wrong.
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not read the sent card for entry {EntryNo}.", approvalEntryNo);
            return null;
        }
    }

    /// <summary>
    /// Removes the record once the card has been refreshed, so a later event
    /// for the same entry does not try to update a card that now shows an
    /// outcome.
    /// </summary>
    public async Task DeleteAsync(int approvalEntryNo, CancellationToken cancellationToken)
    {
        try
        {
            var table = _tableService.GetTableClient(TableName);
            await table.DeleteEntityAsync(
                Partition, approvalEntryNo.ToString(), ETag.All, cancellationToken);
        }
        catch (RequestFailedException ex) when (ex.Status == 404)
        {
            // Already gone.
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not clear the sent card for entry {EntryNo}.", approvalEntryNo);
        }
    }
}

public sealed record SentCard
{
    public required int ApprovalEntryNo { get; init; }
    public required string ConversationId { get; init; }
    public required string ActivityId { get; init; }
    public required string ServiceUrl { get; init; }
    public string DocumentNo { get; init; } = string.Empty;
}
