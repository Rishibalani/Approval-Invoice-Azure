using System.Text.Json;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Options;
using PulseNet.Approvals.Functions.Security;
using PulseNet.Approvals.Functions.Services;

namespace PulseNet.Approvals.Functions.Functions;

/// <summary>
/// The WhatsApp webhook. Everything Meta sends arrives here.
///
/// TWO METHODS, TWO PURPOSES
///
///   GET   the verification handshake, once, when the URL is saved in the
///         Meta dashboard. Echo the challenge back as PLAIN TEXT.
///   POST  everything after that - button taps, typed messages, delivery
///         receipts.
///
/// ALWAYS RETURN 200
///
/// Meta retries a webhook for up to 24 hours if it does not get one, with
/// increasing frequency. Returning 500 on a payload we cannot handle turns one
/// bad message into a day of retries, and every retry is another chance to
/// double-action an approval.
///
/// So: acknowledge everything, and handle what we recognise. The only thing
/// that gets a non-200 is a failed signature, because that is not Meta.
///
/// THE REJECTION FLOW, WHICH IS UNIQUE TO THIS CHANNEL
///
/// A quick-reply button carries a payload and nothing else - there is no text
/// field on a button. So a rejection reason cannot arrive with the tap:
///
///   1. Approver taps Reject
///   2. We store the pending rejection and ask for a reason as free text
///      (legal, because their tap just opened the 24-hour service window)
///   3. They type it
///   4. We match it to the pending rejection and execute
///
/// Teams and Outlook do this in one interaction. WhatsApp cannot.
/// </summary>
public sealed class WhatsAppWebhookFunction
{
    private readonly WhatsAppSignatureValidator _signatureValidator;
    private readonly ApprovalDecisionService _decisionService;
    private readonly PendingRejectionStore _pendingRejections;
    private readonly IdempotencyStore _idempotencyStore;
    private readonly WhatsAppClient _client;
    private readonly WhatsAppChannelOptions _options;
    private readonly ILogger<WhatsAppWebhookFunction> _logger;

    public WhatsAppWebhookFunction(
        WhatsAppSignatureValidator signatureValidator,
        ApprovalDecisionService decisionService,
        PendingRejectionStore pendingRejections,
        IdempotencyStore idempotencyStore,
        WhatsAppClient client,
        IOptions<ChannelOptions> options,
        ILogger<WhatsAppWebhookFunction> logger)
    {
        _signatureValidator = signatureValidator;
        _decisionService = decisionService;
        _pendingRejections = pendingRejections;
        _idempotencyStore = idempotencyStore;
        _client = client;
        _options = options.Value.WhatsApp;
        _logger = logger;
    }

    // ------------------------------------------------------------------
    //  GET - verification handshake
    // ------------------------------------------------------------------

    [Function(nameof(WhatsAppVerify))]
    public IActionResult WhatsAppVerify(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "whatsapp/webhook")]
        HttpRequest req)
    {
        // Meta sends hub.mode. Some proxies - dev tunnels among them - also
        // deliver an underscored copy, and depending on the pipeline the
        // dotted form can arrive unreadable. Reading both spellings costs
        // nothing and avoids a failure that looks like a token mismatch when
        // the token was simply never read.
        var mode = QueryValue(req, "hub.mode", "hub_mode");
        var token = QueryValue(req, "hub.verify_token", "hub_verify_token");
        var challenge = QueryValue(req, "hub.challenge", "hub_challenge");

        if (!_signatureValidator.IsValidVerification(mode, token))
        {
            _logger.LogWarning(
                "WhatsApp webhook verification refused. mode='{Mode}' tokenPresent={HasToken}. " +
                "If the token looks present, compare it byte for byte with Channels:WhatsApp:WebhookVerifyToken.",
                mode ?? "(absent)",
                !string.IsNullOrWhiteSpace(token));

            // 403 with a body, NOT ForbidResult. ForbidResult asks the
            // authentication stack to challenge the caller, and with no scheme
            // registered it throws - which the Functions host then turns into
            // a 200 with an empty body. Meta reads that as a failed challenge
            // and reports "could not be validated", which sends you looking at
            // the tunnel instead of at this line.
            return new ContentResult
            {
                Content = "verification failed",
                ContentType = "text/plain",
                StatusCode = StatusCodes.Status403Forbidden
            };
        }

        _logger.LogInformation("WhatsApp webhook verified.");

        // PLAIN TEXT, unquoted. Returning JSON, or a quoted string, makes Meta
        // reject the subscription with no explanation of why.
        return new ContentResult
        {
            Content = challenge,
            ContentType = "text/plain",
            StatusCode = StatusCodes.Status200OK
        };
    }

    /// <summary>
    /// Reads a query parameter under any of the given names, first match wins.
    /// </summary>
    private static string? QueryValue(HttpRequest req, params string[] names)
    {
        foreach (var name in names)
        {
            var value = req.Query[name].FirstOrDefault();

            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return null;
    }

    // ------------------------------------------------------------------
    //  POST - everything else
    // ------------------------------------------------------------------

    [Function(nameof(WhatsAppWebhook))]
    public async Task<IActionResult> WhatsAppWebhook(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "whatsapp/webhook")]
        HttpRequest req,
        CancellationToken cancellationToken)
    {
        // Read the body ONCE as text. The signature covers these exact bytes,
        // so deserialising and re-serialising would change whitespace and key
        // order, and the check would fail invisibly.
        string rawBody;
        using (var reader = new StreamReader(req.Body))
        {
            rawBody = await reader.ReadToEndAsync(cancellationToken);
        }

        if (!_signatureValidator.IsValid(rawBody, req.Headers["X-Hub-Signature-256"].FirstOrDefault()))
        {
            // The one case that is not Meta, so the one case that gets a 401.
            return new UnauthorizedResult();
        }

        try
        {
            await ProcessAsync(rawBody, cancellationToken);
        }
        catch (Exception ex)
        {
            // Swallow deliberately. A 500 here earns 24 hours of retries, and
            // every retry is another chance to double-action an approval.
            _logger.LogError(ex, "Processing a WhatsApp webhook threw. Acknowledging anyway.");
        }

        return new OkResult();
    }

    // ------------------------------------------------------------------

    private async Task ProcessAsync(string rawBody, CancellationToken cancellationToken)
    {
        using var doc = JsonDocument.Parse(rawBody);

        if (!doc.RootElement.TryGetProperty("entry", out var entries))
        {
            return;
        }

        foreach (var entry in entries.EnumerateArray())
        {
            if (!entry.TryGetProperty("changes", out var changes))
            {
                continue;
            }

            foreach (var change in changes.EnumerateArray())
            {
                if (!change.TryGetProperty("value", out var value))
                {
                    continue;
                }

                // Delivery and read receipts arrive here too. Useful for
                // diagnosing "they never got it", noise otherwise.
                if (value.TryGetProperty("statuses", out var statuses))
                {
                    LogStatuses(statuses);
                }

                if (!value.TryGetProperty("messages", out var messages))
                {
                    continue;
                }

                foreach (var message in messages.EnumerateArray())
                {
                    await HandleMessageAsync(message, cancellationToken);
                }
            }
        }
    }

    private async Task HandleMessageAsync(JsonElement message, CancellationToken cancellationToken)
    {
        var messageId = Str(message, "id");
        var from = Str(message, "from");
        var type = Str(message, "type");

        if (string.IsNullOrWhiteSpace(messageId) || string.IsNullOrWhiteSpace(from))
        {
            return;
        }

        // Meta retries for up to 24 hours. Without this, a slow response means
        // the same tap arrives twice and approves the same invoice twice.
        var isFirstSighting = await _idempotencyStore.TryClaimInTableAsync(
            messageId, _options.InboundDedupeTable, from, cancellationToken);

        if (!isFirstSighting)
        {
            _logger.LogInformation("Duplicate WhatsApp message {MessageId} discarded.", messageId);
            return;
        }

        // Blue ticks. Cosmetic, but its absence makes the bot feel broken
        // during the seconds before a reply arrives.
        await _client.MarkReadAsync(messageId, cancellationToken);

        switch (type)
        {
            case "button":
                // A quick-reply on a TEMPLATE message. The payload is ours.
                await HandleButtonAsync(
                    from,
                    Str(message.GetProperty("button"), "payload"),
                    cancellationToken);
                break;

            case "interactive":
                // A quick-reply on a non-template interactive message. Same
                // idea, different shape - handled so a future interactive
                // follow-up does not silently stop working.
                if (message.TryGetProperty("interactive", out var interactive) &&
                    interactive.TryGetProperty("button_reply", out var buttonReply))
                {
                    await HandleButtonAsync(from, Str(buttonReply, "id"), cancellationToken);
                }
                break;

            case "text":
                // Possibly a rejection reason, possibly ordinary chatter.
                await HandleTextAsync(
                    from,
                    Str(message.GetProperty("text"), "body"),
                    cancellationToken);
                break;

            default:
                // Images, audio, location, and so on. Not an interaction this
                // bot understands, and saying so beats silence.
                await _client.SendTextAsync(
                    from,
                    "I can only handle the Approve and Reject buttons on an approval message. " +
                    "For anything else, please use Business Central.",
                    cancellationToken);
                break;
        }
    }

    /// <summary>
    /// A button was tapped. The payload is the signed action token.
    /// </summary>
    private async Task HandleButtonAsync(string from, string payload, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(payload))
        {
            return;
        }

        // Peek at the action WITHOUT consuming the nonce. A rejection is not
        // executed yet - it waits for a reason - so burning the nonce here
        // would leave the approver unable to complete what they started.
        var peek = await _decisionService.PeekAsync(payload, cancellationToken);

        if (!peek.IsValid)
        {
            await _client.SendTextAsync(from, peek.Message, cancellationToken);
            return;
        }

        if (peek.Action == ApprovalAction.Approve)
        {
            var outcome = await _decisionService.ExecuteAsync(
                payload,
                assertedIdentity: null,
                requireAssertedIdentity: false,
                channel: "WhatsApp",
                deviceInfo: $"WhatsApp ({MaskPhone(from)})",
                cancellationToken: cancellationToken);

            await _client.SendTextAsync(from, outcome.Message, cancellationToken);
            return;
        }

        // Reject. Park it and ask for a reason.
        await _pendingRejections.SaveAsync(new PendingRejection
        {
            PhoneNumber = from,
            Token = payload,
            ApprovalEntryNo = peek.ApprovalEntryNo,
            DocumentNo = peek.DocumentNo ?? string.Empty
        }, cancellationToken);

        await _client.SendTextAsync(
            from,
            "Please reply with the reason for rejecting this invoice. " +
            $"Your reply within the next {_options.RejectionReasonTimeoutMinutes} minutes will be recorded in Business Central.",
            cancellationToken);
    }

    /// <summary>
    /// Free text. Either the reason for a rejection we are waiting on, or
    /// somebody talking to a bot that does not converse.
    /// </summary>
    private async Task HandleTextAsync(string from, string text, CancellationToken cancellationToken)
    {
        var pending = await _pendingRejections.TakeAsync(from, cancellationToken);

        if (pending is null)
        {
            await _client.SendTextAsync(
                from,
                "There is nothing waiting for a reply. Approval requests arrive with Approve and Reject buttons - " +
                "please use those, or open the invoice in Business Central.",
                cancellationToken);

            return;
        }

        if (string.IsNullOrWhiteSpace(text))
        {
            await _client.SendTextAsync(
                from,
                "A reason is required to reject an invoice. Please reply with a few words explaining why.",
                cancellationToken);

            // Put it back - they tried, the text was empty, and making them
            // start over from the original message would be worse.
            await _pendingRejections.SaveAsync(pending, cancellationToken);
            return;
        }

        var outcome = await _decisionService.ExecuteAsync(
            pending.Token,
            assertedIdentity: null,
            requireAssertedIdentity: false,
            channel: "WhatsApp",
            deviceInfo: $"WhatsApp ({MaskPhone(from)})",
            cancellationToken: cancellationToken,
            comment: text.Trim(),
            requireRejectionReason: true);

        await _client.SendTextAsync(from, outcome.Message, cancellationToken);
    }

    private void LogStatuses(JsonElement statuses)
    {
        foreach (var status in statuses.EnumerateArray())
        {
            var state = Str(status, "status");

            if (state == "failed")
            {
                // The one status worth an error. Everything else is normal
                // delivery telemetry.
                _logger.LogError(
                    "WhatsApp reported a failed delivery for message {MessageId}: {Detail}",
                    Str(status, "id"),
                    status.TryGetProperty("errors", out var errors) ? errors.ToString() : "no detail");
            }
            else
            {
                _logger.LogDebug(
                    "WhatsApp status {Status} for {MessageId}.", state, Str(status, "id"));
            }
        }
    }

    /// <summary>
    /// Last four digits only, for the audit trail. A full mobile number in a
    /// log is personal data nobody needs in order to reconstruct what happened.
    /// </summary>
    private static string MaskPhone(string phone) =>
        phone.Length <= 4 ? "****" : "****" + phone[^4..];

    private static string Str(JsonElement element, string name) =>
        element.TryGetProperty(name, out var v) ? v.GetString() ?? string.Empty : string.Empty;
}