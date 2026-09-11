using Azure;
using Azure.Data.Tables;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Options;

namespace PulseNet.Approvals.Functions.Services;

/// <summary>
/// Remembers how to reach each approver in Teams.
///
/// WHY THIS EXISTS
///
/// A bot cannot start a 1:1 chat with someone out of nowhere. Teams requires
/// either an existing conversation or an installed app. The moment a user adds
/// the app, Teams sends a conversationUpdate activity carrying everything
/// needed to message them later - conversation ID, service URL, tenant. Miss
/// that moment and there is no second chance without Graph.
///
/// So the bot endpoint captures it on every inbound activity, not just on
/// install. Cheap, idempotent, and it self-heals if a row is ever lost.
///
/// Keyed on Entra object ID because that is the identifier shared with
/// Business Central via Graph. The Teams 29: ID is specific to one bot-user
/// pairing and would break if the bot were ever re-registered.
/// </summary>
public sealed class ConversationReferenceStore
{
    private const string Partition = "conv";

    private readonly TableServiceClient _tableService;
    private readonly TeamsBotOptions _options;
    private readonly ILogger<ConversationReferenceStore> _logger;

    public ConversationReferenceStore(
        TableServiceClient tableService,
        IOptions<TeamsBotOptions> options,
        ILogger<ConversationReferenceStore> logger)
    {
        _tableService = tableService;
        _options = options.Value;
        _logger = logger;
    }

    /// <summary>
    /// Stores or refreshes a reference. Upsert rather than insert: the service
    /// URL can change when Microsoft moves a tenant between regions, and a
    /// stale URL fails in a way that looks like a permissions problem.
    /// </summary>
    public async Task SaveAsync(ConversationReference reference, CancellationToken cancellationToken)
    {
        try
        {
            var table = _tableService.GetTableClient(_options.ConversationTable);
            await table.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

            var entity = new TableEntity(Partition, reference.AadObjectId)
            {
                ["ConversationId"] = reference.ConversationId,
                ["ServiceUrl"] = reference.ServiceUrl,
                ["TenantId"] = reference.TenantId,
                ["UserPrincipalName"] = reference.UserPrincipalName ?? string.Empty,
                ["DisplayName"] = reference.DisplayName ?? string.Empty,
                ["CapturedUtc"] = DateTime.UtcNow
            };

            await table.UpsertEntityAsync(entity, TableUpdateMode.Replace, cancellationToken);

            _logger.LogInformation(
                "Conversation reference stored for {Upn}.",
                reference.UserPrincipalName ?? reference.AadObjectId);
        }
        catch (Exception ex)
        {
            // Never fail an inbound activity because bookkeeping failed. The
            // next activity from this user will try again.
            _logger.LogWarning(ex, "Could not store the conversation reference.");
        }
    }

    public async Task<ConversationReference?> GetAsync(string aadObjectId, CancellationToken cancellationToken)
    {
        try
        {
            var table = _tableService.GetTableClient(_options.ConversationTable);
            await table.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

            var entity = await table.GetEntityAsync<TableEntity>(
                Partition, aadObjectId, cancellationToken: cancellationToken);

            return new ConversationReference
            {
                AadObjectId = aadObjectId,
                ConversationId = entity.Value.GetString("ConversationId") ?? string.Empty,
                ServiceUrl = entity.Value.GetString("ServiceUrl") ?? _options.DefaultServiceUrl,
                TenantId = entity.Value.GetString("TenantId") ?? _options.TenantId,
                UserPrincipalName = entity.Value.GetString("UserPrincipalName"),
                DisplayName = entity.Value.GetString("DisplayName")
            };
        }
        catch (RequestFailedException ex) when (ex.Status == 404)
        {
            // Expected for anyone who has never opened the app. The sender
            // then tries to create a conversation, then falls back.
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not read the conversation reference for {Id}.", aadObjectId);
            return null;
        }
    }

    /// <summary>
    /// Removes a reference when a user uninstalls the app. Keeping a dead one
    /// means every future dispatch spends a failed round trip before falling
    /// back.
    /// </summary>
    public async Task DeleteAsync(string aadObjectId, CancellationToken cancellationToken)
    {
        try
        {
            var table = _tableService.GetTableClient(_options.ConversationTable);
            await table.DeleteEntityAsync(Partition, aadObjectId, ETag.All, cancellationToken);
            _logger.LogInformation("Conversation reference removed for {Id}.", aadObjectId);
        }
        catch (RequestFailedException ex) when (ex.Status == 404)
        {
            // Already gone.
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not remove the conversation reference.");
        }
    }
}
