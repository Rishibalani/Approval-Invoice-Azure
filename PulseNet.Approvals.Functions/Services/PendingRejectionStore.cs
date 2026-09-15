using Azure;
using Azure.Data.Tables;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Options;

namespace PulseNet.Approvals.Functions.Services;

/// <summary>
/// Holds a rejection between the tap and the reason.
///
/// WHY ONLY WHATSAPP NEEDS THIS
///
/// Teams and Outlook can collect a comment in the same interaction as the
/// decision - Action.Execute posts a ShowCard's inputs, Action.Http posts a
/// form body. WhatsApp quick-reply buttons carry a fixed payload and nothing
/// else. There is no text field on a button.
///
/// So rejection becomes two messages: the tap, then the reason typed as free
/// text. Between them something has to remember which invoice is being
/// rejected, and by whom.
///
/// KEYED ON PHONE NUMBER, NOT ON THE APPROVAL
///
/// The reason arrives as an ordinary text message with nothing tying it to an
/// invoice - WhatsApp does not thread replies to buttons. The only correlation
/// available is "this person tapped Reject a moment ago", so the phone number
/// is the key.
///
/// One consequence worth knowing: an approver who taps Reject on two invoices
/// before typing anything would have the second overwrite the first. That is
/// deliberate. Queueing them would mean asking "which one did you mean?",
/// which is worse than a short window where the most recent tap wins.
///
/// EXPIRY IS SHORT ON PURPOSE
///
/// The approver has just tapped a button, so they are present and typing.
/// Fifteen minutes is generous for that, and short enough that a forgotten tap
/// does not leave a rejection armed for half an hour ready to consume an
/// unrelated message.
/// </summary>
public sealed class PendingRejectionStore
{
    private const string Partition = "pending";

    private readonly TableServiceClient _tableService;
    private readonly WhatsAppChannelOptions _options;
    private readonly ILogger<PendingRejectionStore> _logger;

    public PendingRejectionStore(
        TableServiceClient tableService,
        IOptions<ChannelOptions> options,
        ILogger<PendingRejectionStore> logger)
    {
        _tableService = tableService;
        _options = options.Value.WhatsApp;
        _logger = logger;
    }

    public async Task SaveAsync(PendingRejection pending, CancellationToken cancellationToken)
    {
        try
        {
            var table = _tableService.GetTableClient(_options.PendingRejectionTable);
            await table.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

            var entity = new TableEntity(Partition, Normalise(pending.PhoneNumber))
            {
                ["Token"] = pending.Token,
                ["ApprovalEntryNo"] = pending.ApprovalEntryNo,
                ["DocumentNo"] = pending.DocumentNo,
                ["CreatedUtc"] = DateTime.UtcNow
            };

            // Upsert, so a second tap replaces the first rather than failing.
            await table.UpsertEntityAsync(entity, TableUpdateMode.Replace, cancellationToken);
        }
        catch (Exception ex)
        {
            // If this fails the approver gets asked for a reason and then told
            // it expired, which is confusing but harmless - nothing was
            // approved or rejected.
            _logger.LogError(ex, "Could not store the pending rejection.");
        }
    }

    /// <summary>
    /// Retrieves and REMOVES a pending rejection.
    ///
    /// Removal is part of the read, not a separate step. A reason consumed
    /// twice would execute the rejection twice, and while Business Central
    /// would catch the second with ALREADY_PROCESSED, relying on that means
    /// the approver sees an error for something they did once.
    /// </summary>
    public async Task<PendingRejection?> TakeAsync(string phoneNumber, CancellationToken cancellationToken)
    {
        var key = Normalise(phoneNumber);

        try
        {
            var table = _tableService.GetTableClient(_options.PendingRejectionTable);
            await table.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

            var response = await table.GetEntityAsync<TableEntity>(
                Partition, key, cancellationToken: cancellationToken);

            var entity = response.Value;
            var createdUtc = entity.GetDateTime("CreatedUtc") ?? DateTime.MinValue;

            // Delete first, whether or not it turns out to be expired. A stale
            // row left behind would be consumed by the next unrelated message.
            await table.DeleteEntityAsync(Partition, key, ETag.All, cancellationToken);

            if (DateTime.UtcNow - createdUtc > TimeSpan.FromMinutes(_options.RejectionReasonTimeoutMinutes))
            {
                _logger.LogInformation("Pending rejection for {Phone} had expired.", key);
                return null;
            }

            return new PendingRejection
            {
                PhoneNumber = key,
                Token = entity.GetString("Token") ?? string.Empty,
                ApprovalEntryNo = entity.GetInt32("ApprovalEntryNo") ?? 0,
                DocumentNo = entity.GetString("DocumentNo") ?? string.Empty
            };
        }
        catch (RequestFailedException ex) when (ex.Status == 404)
        {
            // Nothing pending. The message was ordinary chatter, not a reason.
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Could not read the pending rejection for {Phone}.", key);
            return null;
        }
    }

    private static string Normalise(string phone) =>
        new(phone.Where(char.IsDigit).ToArray());
}

public sealed record PendingRejection
{
    public required string PhoneNumber { get; init; }
    public required string Token { get; init; }
    public required int ApprovalEntryNo { get; init; }
    public required string DocumentNo { get; init; }
}
