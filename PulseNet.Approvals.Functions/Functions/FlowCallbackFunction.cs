using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Options;
using PulseNet.Approvals.Functions.Services;

namespace PulseNet.Approvals.Functions.Functions;

/// <summary>
/// OPTION D. Receives a button press that a Power Automate flow collected from
/// an Adaptive Card, and runs it through the same decision path as every other
/// channel.
///
/// WHY THE FLOW IS NOT TRUSTED
///
/// The flow knows who responded, but it is an automation anyone with edit
/// rights on it could change. So nothing here takes the flow's word for the
/// decision: the signed token in the press says which approval entry, which
/// approver and which action, and ApprovalDecisionService re-checks all three
/// before Business Central is asked to do anything. The flow's own responder
/// email is used as a SECOND check, not the first.
///
/// TWO LOCKS ON THE DOOR
///
///   1. The Functions host key (AuthorizationLevel.Function).
///   2. A shared secret in x-pn-flow-secret, compared in constant time.
///
/// Either alone would do at a push; both means a leaked URL in a flow-run
/// history is not enough on its own.
///
/// WHAT THE FLOW SHOULD DO WITH THE ANSWER
///
/// The response body carries a title and a message written for an approver.
/// The flow posts it back into the chat, so the person who pressed sees what
/// happened - including "already handled", which is a normal answer and not an
/// error.
/// </summary>
public sealed class FlowCallbackFunction
{
    private readonly ApprovalDecisionService _decisionService;
    private readonly ChannelOptions _channelOptions;
    private readonly ILogger<FlowCallbackFunction> _logger;

    public FlowCallbackFunction(
        ApprovalDecisionService decisionService,
        IOptions<ChannelOptions> channelOptions,
        ILogger<FlowCallbackFunction> logger)
    {
        _decisionService = decisionService;
        _channelOptions = channelOptions.Value;
        _logger = logger;
    }

    [Function(nameof(FlowCallback))]
    public async Task<IActionResult> FlowCallback(
        [HttpTrigger(AuthorizationLevel.Function, "post", Route = "approvals/flow-callback")]
        HttpRequest req,
        CancellationToken cancellationToken)
    {
        var correlationId = Guid.NewGuid().ToString();

        // ---- Lock 2: the shared secret ---------------------------------
        if (!SecretMatches(req.Headers[FlowContract.SecretHeader].FirstOrDefault()))
        {
            _logger.LogWarning("Flow callback rejected: bad or missing x-pn-flow-secret.");
            return new UnauthorizedResult();
        }

        // ---- The press -------------------------------------------------
        string body;

        using (var reader = new StreamReader(req.Body, Encoding.UTF8))
        {
            body = await reader.ReadToEndAsync(cancellationToken);
        }

        FlowCallbackRequest? press;

        try
        {
            press = JsonSerializer.Deserialize<FlowCallbackRequest>(body, JsonOptions);
        }
        catch (JsonException ex)
        {
            _logger.LogWarning(ex, "Flow callback body was not valid JSON.");
            return Answer(StatusCodes.Status400BadRequest, "Could not read that", "The response could not be read. Please use Business Central.");
        }

        if (press is null || string.IsNullOrWhiteSpace(press.ActionToken))
        {
            _logger.LogWarning("Flow callback carried no token.");
            return Answer(StatusCodes.Status400BadRequest, "Could not read that", "This card is missing its approval reference. Please use Business Central.");
        }

        // ---- The decision ----------------------------------------------
        // requireAssertedIdentity: true. The flow always knows who responded -
        // Teams told it - so a missing responder means something is wrong with
        // the flow, not with the approver, and waving it through would leave an
        // approval nobody can be held to.
        var outcome = await _decisionService.ExecuteAsync(
            press.ActionToken,
            press.ResponderEmail,
            requireAssertedIdentity: true,
            channel: "Teams",
            deviceInfo: $"Power Automate flow ({press.ResponderEmail ?? "unknown responder"})",
            cancellationToken: cancellationToken,
            comment: press.Comment,
            requireRejectionReason: _channelOptions.Outlook.RequireRejectionReason);

        _logger.LogInformation(
            "Flow callback for {DocumentNo} by {Responder}: {Title}. ref {CorrelationId}",
            press.DocumentNo ?? "(unknown document)",
            press.ResponderEmail ?? "(unasserted)",
            outcome.Title,
            press.CorrelationId ?? correlationId);

        // 200 whatever the business outcome. A refusal is an answer, not a
        // transport failure: a non-2xx would make Power Automate retry a press
        // that was already handled correctly.
        return Answer(StatusCodes.Status200OK, outcome.Title, outcome.Message, outcome.Success);
    }

    // ------------------------------------------------------------------

    private bool SecretMatches(string? presented)
    {
        var expected = _channelOptions.Teams.FlowCallbackSecret;

        if (string.IsNullOrWhiteSpace(expected))
        {
            _logger.LogError(
                "Channels__Teams__FlowCallbackSecret is not set, so no flow callback can be accepted.");

            return false;
        }

        if (string.IsNullOrWhiteSpace(presented))
        {
            return false;
        }

        // Constant time. A byte-by-byte comparison that returns early leaks the
        // secret one character at a time to anyone who can measure it.
        return CryptographicOperations.FixedTimeEquals(
            Encoding.UTF8.GetBytes(presented),
            Encoding.UTF8.GetBytes(expected));
    }

    private static ObjectResult Answer(int statusCode, string title, string message, bool succeeded = false) =>
        new(new { title, message, succeeded }) { StatusCode = statusCode };

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    /// <summary>
    /// What the flow posts back. The token, `pnAction` and the ids come from
    /// the card's own action data; `comment` is what the approver typed;
    /// `responderEmail` is who Teams says pressed the button.
    ///
    /// THE TOKEN ARRIVES AS pnToken
    ///
    /// ApprovalCardBuilder writes the signed token into the Action.Submit data
    /// as `pnToken`, so a flow that hands the card response straight back sends
    /// that name. `token` is accepted as well, for a flow that renames it while
    /// building the request body. Read both through <see cref="ActionToken"/>
    /// rather than either property directly.
    /// </summary>
    public sealed record FlowCallbackRequest
    {
        /// <summary>The name ApprovalCardBuilder actually puts on the card.</summary>
        [JsonPropertyName("pnToken")] public string? PnToken { get; init; }

        /// <summary>Accepted alias, for a flow that maps the field itself.</summary>
        [JsonPropertyName("token")] public string? Token { get; init; }

        /// <summary>Whichever of the two the flow sent.</summary>
        [JsonIgnore]
        public string? ActionToken =>
            !string.IsNullOrWhiteSpace(PnToken) ? PnToken : Token;

        public string? Comment { get; init; }
        public string? ResponderEmail { get; init; }
        public string? DocumentNo { get; init; }
        public string? CorrelationId { get; init; }
    }
}
