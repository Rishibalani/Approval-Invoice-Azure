using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Options;
using PulseNet.Approvals.Functions.Security;
using PulseNet.Approvals.Functions.Services;

namespace PulseNet.Approvals.Functions.Functions;

/// <summary>
/// Receives the Action.Http POST when someone taps Approve or Reject inside
/// Outlook.
///
/// TWO RESPONSE HEADERS DO ALL THE WORK
///
///   CARD-ACTION-STATUS   a short message Outlook shows the user in place of
///                        the buttons. This is the entire user feedback
///                        mechanism - there is no page to redirect to.
///
///   CARD-UPDATE-IN-BODY  set to true, with a replacement card in the response
///                        body, and Outlook swaps the card in the email
///                        itself. That is how a decided approval stops looking
///                        live in the inbox.
///
/// Get the headers wrong and the approval still succeeds, but the user sees a
/// generic failure and taps again. Worth testing explicitly.
///
/// AUTHENTICATION
///
/// Anonymous at the Function level, because Outlook cannot attach a function
/// key. The Authorization header carries a Microsoft-issued token which
/// ActionableMessageTokenValidator checks, and our own signed action token
/// arrives in the body.
///
/// Both are needed and neither substitutes for the other: the Microsoft token
/// proves the request came from Outlook, the action token proves what the
/// approver was shown and that it has not already been used.
/// </summary>
public sealed class OutlookActionFunction
{
    private readonly ActionableMessageTokenValidator _tokenValidator;
    private readonly ApprovalDecisionService _decisionService;
    private readonly ChannelOptions _options;
    private readonly ILogger<OutlookActionFunction> _logger;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public OutlookActionFunction(
        ActionableMessageTokenValidator tokenValidator,
        ApprovalDecisionService decisionService,
        IOptions<ChannelOptions> options,
        ILogger<OutlookActionFunction> logger)
    {
        _tokenValidator = tokenValidator;
        _decisionService = decisionService;
        _options = options.Value;
        _logger = logger;
    }

    [Function(nameof(OutlookAction))]
    public async Task<IActionResult> OutlookAction(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "approvals/outlook")]
        HttpRequest req,
        CancellationToken cancellationToken)
    {
        string rawBody;
        using (var reader = new StreamReader(req.Body))
        {
            rawBody = await reader.ReadToEndAsync(cancellationToken);
        }

        // ---- 1. Did this really come from Outlook? --------------------
        var validation = await _tokenValidator.ValidateAsync(
            req.Headers.Authorization.FirstOrDefault(),
            cancellationToken);

        if (!validation.IsValid)
        {
            _logger.LogWarning("Rejected an Outlook action: {Reason}", validation.FailureReason);

            // No detail in the response. Telling an unauthenticated caller
            // which check failed helps them get past the next one.
            return Status(
                StatusCodes.Status401Unauthorized,
                "This action could not be verified. Please use Business Central.");
        }

        // ---- 2. What are we being asked to do? ------------------------
        OutlookActionRequest? request;
        try
        {
            request = JsonSerializer.Deserialize<OutlookActionRequest>(rawBody, JsonOptions);
        }
        catch (JsonException ex)
        {
            _logger.LogError(ex, "Malformed Outlook action body.");
            return Status(StatusCodes.Status400BadRequest, "This action could not be read.");
        }

        if (request is null || string.IsNullOrWhiteSpace(request.Token))
        {
            return Status(StatusCodes.Status400BadRequest, "This action is missing its approval reference.");
        }

        // ---- 3. Shared decision path ----------------------------------
        //
        // The Outlook token asserts the MAILBOX the action came from, which is
        // a genuine identity signal - weaker than an interactive sign-in, but
        // strong enough to check against the approver the token was minted for.
        var outcome = await _decisionService.ExecuteAsync(
            request.Token,
            assertedIdentity: validation.SenderEmail,
            requireAssertedIdentity: _options.Outlook.RequireMailboxMatch,
            channel: "Outlook",
            deviceInfo: "Outlook actionable message",
            cancellationToken: cancellationToken);

        if (!string.IsNullOrWhiteSpace(request.Comment))
        {
            // Not yet written to Business Central - the action handler would
            // need a comment parameter. Logged so it is not silently lost, and
            // so the gap is visible rather than assumed handled.
            _logger.LogInformation(
                "Approver comment on entry {EntryNo}: {Comment}",
                outcome.ApprovalEntryNo, request.Comment);
        }

        // ---- 4. Replace the card in the inbox -------------------------
        return CardUpdate(outcome, request.Verb);
    }

    /// <summary>
    /// Replaces the card in the email with a record of the outcome, and shows
    /// a short status line while it happens.
    /// </summary>
    private IActionResult CardUpdate(ApprovalDecisionOutcome outcome, string? verb)
    {
        var card = new JsonObject
        {
            ["type"] = "AdaptiveCard",
            ["$schema"] = "http://adaptivecards.io/schemas/adaptive-card.json",
            ["version"] = "1.0",
            ["originator"] = _options.Outlook.OriginatorId,
            ["hideOriginalBody"] = true,
            ["body"] = new JsonArray
            {
                new JsonObject
                {
                    ["type"] = "TextBlock",
                    ["text"] = outcome.Title,
                    ["weight"] = "bolder",
                    ["size"] = "medium",
                    ["color"] = outcome.Success ? "good" : "attention",
                    ["wrap"] = true
                },
                new JsonObject
                {
                    ["type"] = "TextBlock",
                    ["text"] = outcome.Message,
                    ["wrap"] = true,
                    ["isSubtle"] = true,
                    ["spacing"] = "small"
                },
                new JsonObject
                {
                    ["type"] = "TextBlock",
                    ["text"] = $"Recorded at {DateTime.UtcNow:yyyy-MM-dd HH:mm} UTC",
                    ["size"] = "small",
                    ["isSubtle"] = true,
                    ["spacing"] = "small"
                }
            }
        };

        var result = new ContentResult
        {
            Content = card.ToJsonString(),
            ContentType = "application/json",
            StatusCode = StatusCodes.Status200OK
        };

        return new OutlookCardResult(result, outcome.Title, updateCard: true);
    }

    private static IActionResult Status(int statusCode, string message) =>
        new OutlookCardResult(
            new StatusCodeResult(statusCode),
            message,
            updateCard: false);
}

/// <summary>
/// The request body Outlook posts, matching the template in OutlookCardBuilder.
/// </summary>
public sealed record OutlookActionRequest
{
    public string? Token { get; init; }
    public string? Verb { get; init; }
    public string? Comment { get; init; }
}

/// <summary>
/// Wraps a result so the two Actionable Message response headers are always
/// set together.
///
/// Separated out because forgetting CARD-ACTION-STATUS is easy and the symptom
/// is misleading: the approval succeeds, Outlook shows a generic failure, and
/// the approver taps again.
/// </summary>
public sealed class OutlookCardResult : IActionResult
{
    private readonly IActionResult _inner;
    private readonly string _status;
    private readonly bool _updateCard;

    public OutlookCardResult(IActionResult inner, string status, bool updateCard)
    {
        _inner = inner;
        _status = status;
        _updateCard = updateCard;
    }

    public async Task ExecuteResultAsync(ActionContext context)
    {
        // Outlook caps what it will display. Truncating here beats having it
        // silently drop a long message.
        context.HttpContext.Response.Headers["CARD-ACTION-STATUS"] =
            _status.Length > 100 ? _status[..100] : _status;

        if (_updateCard)
        {
            context.HttpContext.Response.Headers["CARD-UPDATE-IN-BODY"] = "true";
        }

        await _inner.ExecuteResultAsync(context);
    }
}
