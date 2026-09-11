using System.Text.Json;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Azure.Storage.Queues;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Options;
using PulseNet.Approvals.Functions.Security;
using PulseNet.Approvals.Functions.Services;

namespace PulseNet.Approvals.Functions.Functions;

/// <summary>
/// PHASE 1 ENDPOINT. Everything Business Central talks to.
///
/// Its job is to accept a payload as fast as it safely can, and to be
/// completely certain before it does. It validates, claims the event, drops it
/// on a queue and returns 202. It does not talk to Teams, Graph or anything
/// else - a channel outage must never turn into a Business Central retry storm.
///
/// PHASE 2 picks up from the queue in ApprovalDispatchWorker.
///
/// STATUS CODES AND WHAT BUSINESS CENTRAL DOES WITH THEM
///   202  accepted, queued            -> outbox row moves to Sent
///   409  duplicate idempotency key   -> also treated as Sent, deliberately
///   400  bad payload or wrong schema -> retried, then Failed and alerted
///   401  signature or auth failure   -> retried, then Failed and alerted
///   5xx  our problem                 -> retried with backoff
///
/// The 409 case matters: it is how a retry after a timeout resolves cleanly
/// instead of producing a second notification.
/// </summary>
public sealed class ApprovalIngestFunction
{
    private readonly RequestSignatureValidator _signatureValidator;
    private readonly IdempotencyStore _idempotencyStore;
    private readonly QueueServiceClient _queueService;
    private readonly DispatchOptions _options;
    private readonly ILogger<ApprovalIngestFunction> _logger;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public ApprovalIngestFunction(
        RequestSignatureValidator signatureValidator,
        IdempotencyStore idempotencyStore,
        QueueServiceClient queueService,
        IOptions<DispatchOptions> options,
        ILogger<ApprovalIngestFunction> logger)
    {
        _signatureValidator = signatureValidator;
        _idempotencyStore = idempotencyStore;
        _queueService = queueService;
        _options = options.Value;
        _logger = logger;
    }

    [Function(nameof(ApprovalIngest))]
    public async Task<IActionResult> ApprovalIngest(
        [HttpTrigger(AuthorizationLevel.Function, "post", Route = "approvals/dispatch")]
        HttpRequest req,
        CancellationToken cancellationToken)
    {
        var correlationId = req.Headers["x-pn-correlation-id"].FirstOrDefault() ?? Guid.NewGuid().ToString();

        using var scope = _logger.BeginScope(new Dictionary<string, object>
        {
            ["CorrelationId"] = correlationId
        });

        // ---- Read the body ONCE, as text ---------------------------------
        // The signature is computed over these exact characters. Deserialising
        // first and re-serialising would change whitespace and property order,
        // and the check would fail for reasons nobody can see.
        string rawBody;
        using (var reader = new StreamReader(req.Body))
        {
            rawBody = await reader.ReadToEndAsync(cancellationToken);
        }

        if (string.IsNullOrWhiteSpace(rawBody))
        {
            return Problem(StatusCodes.Status400BadRequest, "empty_body", correlationId);
        }

        // ---- Authenticity ------------------------------------------------
        var signatureResult = await _signatureValidator.ValidateAsync(
            rawBody,
            req.Headers["x-pn-timestamp"].FirstOrDefault(),
            req.Headers["x-pn-nonce"].FirstOrDefault(),
            req.Headers["x-pn-signature"].FirstOrDefault(),
            cancellationToken);

        if (!signatureResult.IsValid)
        {
            // No detail in the response. Telling an unauthenticated caller
            // which check failed helps them get past the next one.
            _logger.LogWarning("Rejected dispatch: {Reason}", signatureResult.FailureReason);
            return Problem(StatusCodes.Status401Unauthorized, "unauthorized", correlationId);
        }

        // ---- Shape -------------------------------------------------------
        ApprovalDispatchPayload? payload;
        try
        {
            payload = JsonSerializer.Deserialize<ApprovalDispatchPayload>(rawBody, JsonOptions);
        }
        catch (JsonException ex)
        {
            _logger.LogError(ex, "Malformed dispatch payload.");
            return Problem(StatusCodes.Status400BadRequest, "malformed_json", correlationId);
        }

        if (payload is null)
        {
            return Problem(StatusCodes.Status400BadRequest, "null_payload", correlationId);
        }

        if (!_options.SchemaAllowList.Contains(payload.SchemaVersion))
        {
            // Refuse rather than guess. A field misread as an amount is worse
            // than a refused request.
            _logger.LogError("Unsupported schema version {Version}.", payload.SchemaVersion);
            return Problem(StatusCodes.Status400BadRequest, "unsupported_schema_version", correlationId);
        }

        // ---- Environment and tenant isolation ----------------------------
        // Cheapest possible guard against a sandbox raising a real approval.
        if (_options.EnvironmentAllowList.Count > 0 &&
            !_options.EnvironmentAllowList.Contains(payload.Source.Environment))
        {
            _logger.LogWarning(
                "Rejected payload from environment {Environment}, which is not on the allow list.",
                payload.Source.Environment);
            return Problem(StatusCodes.Status403Forbidden, "environment_not_allowed", correlationId);
        }

        if (_options.TenantAllowList.Count > 0 &&
            !_options.TenantAllowList.Contains(payload.Source.TenantId))
        {
            _logger.LogWarning("Rejected payload from unknown tenant.");
            return Problem(StatusCodes.Status403Forbidden, "tenant_not_allowed", correlationId);
        }

        // ---- Sanity checks on the business content -----------------------
        if (payload.Approval.ApprovalEntryNo <= 0)
        {
            return Problem(StatusCodes.Status400BadRequest, "missing_approval_entry", correlationId);
        }

        if (string.IsNullOrWhiteSpace(payload.EventId))
        {
            return Problem(StatusCodes.Status400BadRequest, "missing_event_id", correlationId);
        }

        // ---- Idempotency -------------------------------------------------
        var claimed = await _idempotencyStore.TryClaimAsync(
            payload.EventId,
            payload.Source.TenantId,
            correlationId,
            cancellationToken);

        if (!claimed)
        {
            // 409 is a success from Business Central's point of view. Saying so
            // clearly here is what stops a timeout producing two cards.
            return new ObjectResult(new
            {
                status = "duplicate",
                message = "This event was already accepted.",
                eventId = payload.EventId,
                correlationId
            })
            { StatusCode = StatusCodes.Status409Conflict };
        }

        // ---- Hand off ----------------------------------------------------
        _logger.LogInformation(
            "Accepted {EventType} for {DocumentType} {DocumentNo}, approver {Approver}, amount {Amount} {Currency}, canApproveInChannel={CanApprove}",
            payload.EventType,
            payload.Document.DocumentType,
            payload.Document.DocumentNo,
            payload.Approver.UserId,
            payload.Document.AmountLcy,
            payload.Document.CurrencyCode,
            payload.Policy.CanApproveInChannel);

        // In phase 1 the queue message IS the deliverable: it proves the payload
        // arrived intact and is durably parked. Phase 2 adds the worker that
        // reads it and renders a card.
        //
        // The queue is written with an explicit client rather than an output
        // binding because this function has several early-return paths and an
        // output binding would fire on all of them.
        var queue = _queueService.GetQueueClient(_options.QueueName);
        await queue.CreateIfNotExistsAsync(cancellationToken: cancellationToken);
        // No manual base64 here: the client is configured with
        // MessageEncoding.Base64 in Program.cs, which is what the queue trigger
        // in the worker expects. Encoding twice is a classic silent corruption.
        await queue.SendMessageAsync(rawBody, cancellationToken);

        return new ObjectResult(new
        {
            status = "accepted",
            eventId = payload.EventId,
            correlationId,
            queuedUtc = DateTime.UtcNow
        })
        { StatusCode = StatusCodes.Status202Accepted };
    }

    /// <summary>
    /// Health probe used by the Test Connection action on the Business Central
    /// setup page. Anonymous by design so a misconfigured key still produces a
    /// useful answer; it returns nothing sensitive.
    /// </summary>
    [Function(nameof(Health))]
    public IActionResult Health(
        [HttpTrigger(AuthorizationLevel.Function, "get", Route = "health")] HttpRequest req)
    {
        return new OkObjectResult(new
        {
            status = "healthy",
            utc = DateTime.UtcNow,
            schemaVersions = _options.SchemaAllowList,
            signingConfigured = !string.IsNullOrWhiteSpace(_options.SigningSecret)
        });
    }

    private static ObjectResult Problem(int statusCode, string code, string correlationId) =>
        new(new { error = code, correlationId }) { StatusCode = statusCode };
}
