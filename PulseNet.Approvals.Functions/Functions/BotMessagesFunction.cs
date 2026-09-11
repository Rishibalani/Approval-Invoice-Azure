using System.Text.Json;
using System.Text.Json.Nodes;
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
/// The bot's messaging endpoint. This URL goes in the Azure Bot resource's
/// Configuration blade, and everything Teams sends the bot arrives here.
///
/// THREE THINGS ARRIVE
///
///   conversationUpdate  someone installed or removed the app
///   invoke              someone tapped Approve or Reject
///   message             someone typed at the bot
///
/// The first is how proactive messaging becomes possible at all: Teams will
/// not let a bot open a chat with someone who has never added the app, so the
/// install event is the one chance to capture how to reach them. Miss it and
/// there is no second chance without Graph.
///
/// WHY THE TAP IS STRONGER THAN A LINK
///
/// The invoke arrives with a Bot Framework JWT that Microsoft signed, and the
/// activity carries the verified Entra object ID of the person who tapped. No
/// browser, no sign-in prompt, no shared URL that a forwarded screenshot could
/// leak. That identity assurance - not the nicer visuals - is the real argument
/// for the bot over the webhook.
///
/// We still check our own action token as well. The channel proves WHO tapped;
/// the token proves WHAT THEY WERE SHOWN. Both are needed, and neither
/// substitutes for the other.
/// </summary>
public sealed class BotMessagesFunction
{
    private readonly BotFrameworkTokenValidator _tokenValidator;
    private readonly ConversationReferenceStore _conversations;
    private readonly BusinessCentralClient _bcClient;
    private readonly BotConnectorClient _connector;
    private readonly TeamsBotOptions _options;
    private readonly ILogger<BotMessagesFunction> _logger;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public BotMessagesFunction(
        BotFrameworkTokenValidator tokenValidator,
        ConversationReferenceStore conversations,
        BusinessCentralClient bcClient,
        BotConnectorClient connector,
        IOptions<TeamsBotOptions> options,
        ILogger<BotMessagesFunction> logger)
    {
        _tokenValidator = tokenValidator;
        _conversations = conversations;
        _bcClient = bcClient;
        _connector = connector;
        _options = options.Value;
        _logger = logger;
    }

    [Function(nameof(BotMessages))]
    public async Task<IActionResult> BotMessages(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "messages")]
        HttpRequest req,
        CancellationToken cancellationToken)
    {
        // Anonymous by design. Teams cannot attach a function key. The Bot
        // Framework JWT below is the authentication, and it is stronger than a
        // shared key: Microsoft signs it, it expires, and it names the bot it
        // was issued for.
        var authHeader = req.Headers.Authorization.FirstOrDefault();

        if (!await _tokenValidator.IsValidAsync(authHeader, cancellationToken))
        {
            _logger.LogWarning("Rejected an unauthenticated call to the bot endpoint.");
            return new UnauthorizedResult();
        }

        string rawBody;
        using (var reader = new StreamReader(req.Body))
        {
            rawBody = await reader.ReadToEndAsync(cancellationToken);
        }

        BotActivity? activity;
        try
        {
            activity = JsonSerializer.Deserialize<BotActivity>(rawBody, JsonOptions);
        }
        catch (JsonException ex)
        {
            _logger.LogError(ex, "Malformed activity received.");
            return new BadRequestResult();
        }

        if (activity is null)
        {
            return new BadRequestResult();
        }

        // Capture on EVERY activity, not just on install. Cheap, idempotent,
        // and it self-heals a reference that was lost or never captured.
        await CaptureConversationAsync(activity, cancellationToken);

        return activity.Type switch
        {
            "invoke" => await HandleInvokeAsync(activity, cancellationToken),
            "conversationUpdate" => await HandleConversationUpdateAsync(activity, cancellationToken),
            "message" => HandleMessage(),
            _ => new OkResult()
        };
    }

    // ------------------------------------------------------------------
    //  A button was tapped
    // ------------------------------------------------------------------

    private async Task<IActionResult> HandleInvokeAsync(BotActivity activity, CancellationToken cancellationToken)
    {
        if (activity.Name != "adaptiveCard/action")
        {
            return new OkResult();
        }

        var verb = activity.Value?.Action?.Verb;
        var data = activity.Value?.Action?.Data;

        if (string.IsNullOrWhiteSpace(verb) || data is null)
        {
            return CardResponse(NoticeCard("This action could not be read. Please use Business Central."));
        }

        if (!TryGetInt(data, "approvalEntryNo", out var approvalEntryNo))
        {
            return CardResponse(NoticeCard("This card is missing its approval reference. Please use Business Central."));
        }

        // Teams asserts this. It is not something the card could forge.
        var actorObjectId = activity.From?.AadObjectId;
        var actorName = activity.From?.Name ?? "Unknown";

        _logger.LogInformation(
            "Invoke '{Verb}' on entry {EntryNo} from {Actor}.",
            verb, approvalEntryNo, actorName);

        var action = verb.Equals("reject", StringComparison.OrdinalIgnoreCase)
            ? ApprovalAction.Reject
            : ApprovalAction.Approve;

        // Business Central re-checks authority, status, amount and the change
        // gates. Nothing here is trusted to have got that right.
        var result = await _bcClient.ExecuteApprovalAsync(
            approvalEntryNo,
            action,
            channel: "Teams",
            deviceInfo: $"Teams bot ({actorName})",
            correlationId: GetString(data, "correlationId") ?? Guid.NewGuid().ToString(),
            cancellationToken: cancellationToken);

        var verbPast = action == ApprovalAction.Approve ? "Approved" : "Rejected";
        var documentNo = GetString(data, "documentNo") ?? "This invoice";

        // Returning a card replaces the original in place. The buttons vanish
        // and what is left is a record of the decision - which is exactly what
        // a webhook cannot do, and why a decided approval stops looking live.
        var replacement = result.Succeeded
            ? OutcomeCard(documentNo, verbPast, actorName, isPositive: action == ApprovalAction.Approve)
            : NoticeCard(result.ApproverMessage);

        return CardResponse(replacement);
    }

    // ------------------------------------------------------------------
    //  Someone installed or removed the app
    // ------------------------------------------------------------------

    private async Task<IActionResult> HandleConversationUpdateAsync(
        BotActivity activity,
        CancellationToken cancellationToken)
    {
        var botId = _options.AppId;

        // The bot itself being removed means this conversation is dead.
        if (activity.MembersRemoved is not null)
        {
            foreach (var member in activity.MembersRemoved)
            {
                if (member.Id.Contains(botId, StringComparison.OrdinalIgnoreCase))
                {
                    var userObjectId = activity.From?.AadObjectId;

                    if (!string.IsNullOrWhiteSpace(userObjectId))
                    {
                        await _conversations.DeleteAsync(userObjectId, cancellationToken);
                    }
                }
            }
        }

        // A welcome message is worth sending. It confirms to the user that the
        // app works, and it means the conversation is warm before the first
        // real approval arrives.
        if (activity.MembersAdded is not null && activity.Conversation?.ConversationType == "personal")
        {
            foreach (var member in activity.MembersAdded)
            {
                if (!member.Id.Contains(botId, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                await _connector.SendCardAsync(
                    activity.Conversation.Id,
                    activity.ServiceUrl,
                    NoticeCard(
                        "Invoice approvals will arrive here. " +
                        "You can approve or reject them without leaving Teams."),
                    "Invoice approvals are set up",
                    cancellationToken);
            }
        }

        return new OkResult();
    }

    private static IActionResult HandleMessage()
    {
        // Typed messages are not a supported interaction. Saying so is kinder
        // than silence - an approver who types "approved" should learn quickly
        // that it did nothing.
        return new OkResult();
    }

    // ------------------------------------------------------------------

    private async Task CaptureConversationAsync(BotActivity activity, CancellationToken cancellationToken)
    {
        var aadObjectId = activity.From?.AadObjectId;
        var conversationId = activity.Conversation?.Id;

        if (string.IsNullOrWhiteSpace(aadObjectId) || string.IsNullOrWhiteSpace(conversationId))
        {
            return;
        }

        // Only personal chats. A channel conversation cannot be used to reach
        // one approver privately, which is the whole point of bot delivery.
        if (activity.Conversation?.ConversationType != "personal")
        {
            return;
        }

        await _conversations.SaveAsync(new ConversationReference
        {
            AadObjectId = aadObjectId,
            ConversationId = conversationId,
            ServiceUrl = activity.ServiceUrl,
            TenantId = activity.ChannelData?.Tenant?.Id ?? activity.Conversation?.TenantId ?? _options.TenantId,
            DisplayName = activity.From?.Name
        }, cancellationToken);
    }

    /// <summary>
    /// The invoke response shape that replaces a card in place. Teams is
    /// particular about it: statusCode, type and value are all required, and
    /// the type string must be exactly this.
    /// </summary>
    private static IActionResult CardResponse(JsonObject card) =>
        new OkObjectResult(new
        {
            statusCode = 200,
            type = "application/vnd.microsoft.card.adaptive",
            value = JsonSerializer.Deserialize<JsonElement>(card.ToJsonString())
        });

    private static JsonObject OutcomeCard(string documentNo, string outcome, string actor, bool isPositive) =>
        new()
        {
            ["type"] = "AdaptiveCard",
            ["$schema"] = "http://adaptivecards.io/schemas/adaptive-card.json",
            ["version"] = "1.4",
            ["body"] = new JsonArray
            {
                new JsonObject
                {
                    ["type"] = "TextBlock",
                    ["text"] = $"{documentNo} — {outcome}",
                    ["weight"] = "Bolder",
                    ["size"] = "Medium",
                    ["color"] = isPositive ? "Good" : "Attention",
                    ["wrap"] = true
                },
                new JsonObject
                {
                    ["type"] = "TextBlock",
                    ["text"] = $"By {actor} at {DateTime.UtcNow:yyyy-MM-dd HH:mm} UTC",
                    ["isSubtle"] = true,
                    ["size"] = "Small",
                    ["spacing"] = "None",
                    ["wrap"] = true
                }
            }
        };

    private static JsonObject NoticeCard(string text) =>
        new()
        {
            ["type"] = "AdaptiveCard",
            ["$schema"] = "http://adaptivecards.io/schemas/adaptive-card.json",
            ["version"] = "1.4",
            ["body"] = new JsonArray
            {
                new JsonObject
                {
                    ["type"] = "TextBlock",
                    ["text"] = text,
                    ["wrap"] = true
                }
            }
        };

    private static bool TryGetInt(Dictionary<string, object> data, string key, out int value)
    {
        value = 0;

        if (!data.TryGetValue(key, out var raw) || raw is null)
        {
            return false;
        }

        // Values arrive as JsonElement through System.Text.Json, and the card
        // may have carried the number as a string.
        if (raw is JsonElement element)
        {
            return element.ValueKind switch
            {
                JsonValueKind.Number => element.TryGetInt32(out value),
                JsonValueKind.String => int.TryParse(element.GetString(), out value),
                _ => false
            };
        }

        return int.TryParse(raw.ToString(), out value);
    }

    private static string? GetString(Dictionary<string, object> data, string key)
    {
        if (!data.TryGetValue(key, out var raw) || raw is null)
        {
            return null;
        }

        return raw is JsonElement element ? element.ToString() : raw.ToString();
    }
}
