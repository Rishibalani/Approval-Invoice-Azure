using System.Text.Json.Serialization;

namespace PulseNet.Approvals.Functions.Models;

/// <summary>
/// The slice of the Bot Framework Activity schema this solution actually uses.
///
/// Deliberately hand-rolled rather than pulled from Microsoft.Bot.Builder. The
/// full SDK assumes ASP.NET hosting and an adapter pipeline, which fights the
/// isolated Functions worker model. Three activity types and a REST call are
/// all that is needed here, and this way the contract is visible in one file.
/// </summary>
public sealed record BotActivity
{
    /// <summary>message, conversationUpdate, invoke, installationUpdate.</summary>
    [JsonPropertyName("type")] public string Type { get; init; } = string.Empty;

    [JsonPropertyName("id")] public string? Id { get; init; }

    /// <summary>For invoke: "adaptiveCard/action" when a button is tapped.</summary>
    [JsonPropertyName("name")] public string? Name { get; init; }

    /// <summary>
    /// Where to send replies. Always prefer this over configuration - Teams
    /// can move a tenant between regions and the activity is authoritative.
    /// </summary>
    [JsonPropertyName("serviceUrl")] public string ServiceUrl { get; init; } = string.Empty;

    [JsonPropertyName("channelId")] public string? ChannelId { get; init; }
    [JsonPropertyName("from")] public BotAccount? From { get; init; }
    [JsonPropertyName("recipient")] public BotAccount? Recipient { get; init; }
    [JsonPropertyName("conversation")] public BotConversation? Conversation { get; init; }
    [JsonPropertyName("channelData")] public BotChannelData? ChannelData { get; init; }
    [JsonPropertyName("value")] public BotInvokeValue? Value { get; init; }

    [JsonPropertyName("membersAdded")] public IReadOnlyList<BotAccount>? MembersAdded { get; init; }
    [JsonPropertyName("membersRemoved")] public IReadOnlyList<BotAccount>? MembersRemoved { get; init; }

    [JsonPropertyName("replyToId")] public string? ReplyToId { get; init; }
}

public sealed record BotAccount
{
    /// <summary>Teams-scoped ID, e.g. 29:1abc... Not the Entra object ID.</summary>
    [JsonPropertyName("id")] public string Id { get; init; } = string.Empty;

    [JsonPropertyName("name")] public string? Name { get; init; }

    /// <summary>
    /// The Entra object ID. Present on Teams activities and the only identifier
    /// stable enough to key a conversation reference on - the 29: ID is
    /// specific to one bot-user pairing.
    /// </summary>
    [JsonPropertyName("aadObjectId")] public string? AadObjectId { get; init; }
}

public sealed record BotConversation
{
    [JsonPropertyName("id")] public string Id { get; init; } = string.Empty;
    [JsonPropertyName("conversationType")] public string? ConversationType { get; init; }
    [JsonPropertyName("tenantId")] public string? TenantId { get; init; }
    [JsonPropertyName("isGroup")] public bool IsGroup { get; init; }
}

public sealed record BotChannelData
{
    [JsonPropertyName("tenant")] public BotTenant? Tenant { get; init; }
}

public sealed record BotTenant
{
    [JsonPropertyName("id")] public string Id { get; init; } = string.Empty;
}

/// <summary>
/// The payload of an Action.Execute invoke. Teams wraps the card's data object
/// inside action.data.
/// </summary>
public sealed record BotInvokeValue
{
    [JsonPropertyName("action")] public BotInvokeAction? Action { get; init; }
}

public sealed record BotInvokeAction
{
    /// <summary>The verb from the Adaptive Card action - "approve" or "reject".</summary>
    [JsonPropertyName("verb")] public string? Verb { get; init; }

    [JsonPropertyName("data")] public Dictionary<string, object>? Data { get; init; }
}

/// <summary>
/// Everything needed to start a proactive conversation later, captured the
/// first time a user interacts with the bot.
/// </summary>
public sealed record ConversationReference
{
    public required string AadObjectId { get; init; }
    public required string ConversationId { get; init; }
    public required string ServiceUrl { get; init; }
    public required string TenantId { get; init; }
    public string? UserPrincipalName { get; init; }
    public string? DisplayName { get; init; }
    public DateTimeOffset CapturedUtc { get; init; } = DateTimeOffset.UtcNow;
}
