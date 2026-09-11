using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Options;

namespace PulseNet.Approvals.Functions.Services;

/// <summary>
/// Talks to the Bot Connector REST API directly.
///
/// WHY NOT THE BOT FRAMEWORK SDK
///
/// Microsoft.Bot.Builder assumes ASP.NET hosting with an adapter pipeline and
/// a controller. That fights the isolated Functions worker model, and dragging
/// it in would mean adopting its hosting model to use three REST calls.
///
/// Those three calls are: get a token, create a conversation, post an activity.
/// Written directly they fit in one readable file with no hidden behaviour.
///
/// THE CONVERSATION PROBLEM, STATED HONESTLY
///
/// Teams will not let a bot create a 1:1 conversation with someone who has
/// never added the app. CreateConversationAsync returns 403 "bot is not part
/// of the conversation roster". There is no way around this short of
/// proactively installing the app via Graph, which needs
/// TeamsAppInstallation.ReadWriteForUser.All.
///
/// So the sender tries, in order:
///   1. A stored conversation reference, captured when the user installed
///   2. Creating one from the Entra object ID
///   3. Falling back to Outlook
///
/// Step 2 succeeds only if the app is installed but no reference was captured
/// - which happens when the app was installed before this code shipped.
/// </summary>
public sealed class BotConnectorClient
{
    private readonly HttpClient _http;
    private readonly TeamsBotOptions _options;
    private readonly ILogger<BotConnectorClient> _logger;

    private string? _cachedToken;
    private DateTimeOffset _tokenExpiresAt = DateTimeOffset.MinValue;
    private readonly SemaphoreSlim _tokenLock = new(1, 1);

    public BotConnectorClient(
        HttpClient http,
        IOptions<TeamsBotOptions> options,
        ILogger<BotConnectorClient> logger)
    {
        _http = http;
        _options = options.Value;
        _logger = logger;
    }

    // ------------------------------------------------------------------
    //  Conversations
    // ------------------------------------------------------------------

    /// <summary>
    /// Creates a 1:1 conversation with a user, identified by Entra object ID.
    /// Returns the conversation ID, or null if Teams refused - which almost
    /// always means the app is not installed for that user.
    /// </summary>
    public async Task<string?> TryCreateConversationAsync(
        string aadObjectId,
        string serviceUrl,
        CancellationToken cancellationToken)
    {
        var body = new JsonObject
        {
            ["isGroup"] = false,
            ["bot"] = new JsonObject
            {
                ["id"] = $"28:{_options.AppId}"
            },
            ["members"] = new JsonArray
            {
                new JsonObject { ["id"] = aadObjectId }
            },
            ["channelData"] = new JsonObject
            {
                ["tenant"] = new JsonObject { ["id"] = _options.TenantId }
            },
            ["tenantId"] = _options.TenantId
        };

        var url = $"{serviceUrl.TrimEnd('/')}/v3/conversations";

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(body.ToJsonString(), Encoding.UTF8, "application/json")
            };

            request.Headers.Authorization = new AuthenticationHeaderValue(
                "Bearer", await GetTokenAsync(TeamsBotOptions.BotFrameworkScope, cancellationToken));

            using var response = await _http.SendAsync(request, cancellationToken);
            var responseText = await response.Content.ReadAsStringAsync(cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                // 403 here is the expected, uninteresting case: the app is not
                // installed for this user. Logged at Information because it is
                // a routine branch, not a fault.
                if (response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                {
                    _logger.LogInformation(
                        "Cannot open a chat with {AadObjectId} - the Teams app is not installed for them.",
                        aadObjectId);
                }
                else
                {
                    _logger.LogError(
                        "Create conversation failed: {Status} {Body}",
                        (int)response.StatusCode, Truncate(responseText, 400));
                }

                return null;
            }

            using var doc = JsonDocument.Parse(responseText);
            return doc.RootElement.TryGetProperty("id", out var id) ? id.GetString() : null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Create conversation threw for {AadObjectId}.", aadObjectId);
            return null;
        }
    }

    // ------------------------------------------------------------------
    //  Activities
    // ------------------------------------------------------------------

    /// <summary>
    /// Posts an Adaptive Card into a conversation. Returns the activity ID,
    /// which is what makes an in-place update possible later - keep it.
    /// </summary>
    public async Task<string?> SendCardAsync(
        string conversationId,
        string serviceUrl,
        JsonObject card,
        string summary,
        CancellationToken cancellationToken)
    {
        var activity = BuildCardActivity(card, summary);
        var url = $"{serviceUrl.TrimEnd('/')}/v3/conversations/{conversationId}/activities";

        using var request = new HttpRequestMessage(HttpMethod.Post, url)
        {
            Content = new StringContent(activity.ToJsonString(), Encoding.UTF8, "application/json")
        };

        request.Headers.Authorization = new AuthenticationHeaderValue(
            "Bearer", await GetTokenAsync(TeamsBotOptions.BotFrameworkScope, cancellationToken));

        using var response = await _http.SendAsync(request, cancellationToken);
        var responseText = await response.Content.ReadAsStringAsync(cancellationToken);

        if (!response.IsSuccessStatusCode)
        {
            _logger.LogError(
                "Send activity failed: {Status} {Body}",
                (int)response.StatusCode, Truncate(responseText, 400));

            return null;
        }

        using var doc = JsonDocument.Parse(responseText);
        return doc.RootElement.TryGetProperty("id", out var id) ? id.GetString() : null;
    }

    /// <summary>
    /// Replaces a card that is already in the conversation.
    ///
    /// This is the capability a webhook cannot offer, and the main reason the
    /// bot is worth its setup cost: when an approval is decided, the original
    /// card turns into a record of the decision rather than staying live with
    /// buttons that no longer do anything.
    /// </summary>
    public async Task<bool> UpdateCardAsync(
        string conversationId,
        string activityId,
        string serviceUrl,
        JsonObject card,
        string summary,
        CancellationToken cancellationToken)
    {
        var activity = BuildCardActivity(card, summary);
        var url = $"{serviceUrl.TrimEnd('/')}/v3/conversations/{conversationId}/activities/{activityId}";

        using var request = new HttpRequestMessage(HttpMethod.Put, url)
        {
            Content = new StringContent(activity.ToJsonString(), Encoding.UTF8, "application/json")
        };

        request.Headers.Authorization = new AuthenticationHeaderValue(
            "Bearer", await GetTokenAsync(TeamsBotOptions.BotFrameworkScope, cancellationToken));

        using var response = await _http.SendAsync(request, cancellationToken);

        if (!response.IsSuccessStatusCode)
        {
            // Cosmetic failure. The decision is already recorded in Business
            // Central, so this must never surface as an error to the approver.
            _logger.LogWarning(
                "Could not update the card in place: {Status}. The decision itself was unaffected.",
                (int)response.StatusCode);

            return false;
        }

        return true;
    }

    private static JsonObject BuildCardActivity(JsonObject card, string summary) =>
        new()
        {
            ["type"] = "message",
            // Shown in the notification toast and read by screen readers,
            // where the card itself does not render.
            ["summary"] = summary,
            ["attachments"] = new JsonArray
            {
                new JsonObject
                {
                    ["contentType"] = "application/vnd.microsoft.card.adaptive",
                    ["content"] = card.DeepClone()
                }
            }
        };

    // ------------------------------------------------------------------
    //  Graph - UPN to Entra object ID
    // ------------------------------------------------------------------

    /// <summary>
    /// Resolves a UPN to an Entra object ID. Needed because Teams identifies
    /// people by object ID while Business Central knows them by email.
    ///
    /// Returns null rather than throwing when Graph is not configured or the
    /// user is not found - the caller falls back to another channel, which is
    /// a better outcome than an exception nobody sees.
    /// </summary>
    public async Task<string?> TryResolveObjectIdAsync(string upn, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(_options.GraphClientId))
        {
            _logger.LogInformation(
                "Graph is not configured, so {Upn} cannot be resolved to an object ID.", upn);
            return null;
        }

        try
        {
            var url = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(upn)}?$select=id";

            using var request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue(
                "Bearer", await GetTokenAsync(TeamsBotOptions.GraphScope, cancellationToken));

            using var response = await _http.SendAsync(request, cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogWarning(
                    "Graph could not resolve {Upn}: {Status}", upn, (int)response.StatusCode);
                return null;
            }

            var json = await response.Content.ReadAsStringAsync(cancellationToken);
            using var doc = JsonDocument.Parse(json);

            return doc.RootElement.TryGetProperty("id", out var id) ? id.GetString() : null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Graph lookup threw for {Upn}.", upn);
            return null;
        }
    }

    // ------------------------------------------------------------------
    //  Tokens
    // ------------------------------------------------------------------

    /// <summary>
    /// Client credentials against Entra, cached per scope and renewed five
    /// minutes early.
    ///
    /// SingleTenant bots authenticate against their own tenant's endpoint.
    /// MultiTenant bots used login.microsoftonline.com/botframework.com, which
    /// is what most older samples show - but Microsoft retired MultiTenant for
    /// new bot registrations on 31 July 2025, so that path no longer applies.
    /// </summary>
    private async Task<string> GetTokenAsync(string scope, CancellationToken cancellationToken)
    {
        var isGraph = scope == TeamsBotOptions.GraphScope;

        // Only the Bot Framework token is cached on the instance. Graph calls
        // are rare - one per approver, then the object ID is cached in Business
        // Central - so a per-call token is not worth the extra state.
        if (!isGraph && _cachedToken is not null &&
            DateTimeOffset.UtcNow < _tokenExpiresAt.AddMinutes(-5))
        {
            return _cachedToken;
        }

        await _tokenLock.WaitAsync(cancellationToken);
        try
        {
            if (!isGraph && _cachedToken is not null &&
                DateTimeOffset.UtcNow < _tokenExpiresAt.AddMinutes(-5))
            {
                return _cachedToken;
            }

            var clientId = isGraph && !string.IsNullOrWhiteSpace(_options.GraphClientId)
                ? _options.GraphClientId
                : _options.AppId;

            var clientSecret = isGraph && !string.IsNullOrWhiteSpace(_options.GraphClientSecret)
                ? _options.GraphClientSecret
                : _options.AppPassword;

            var form = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["grant_type"] = "client_credentials",
                ["client_id"] = clientId,
                ["client_secret"] = clientSecret,
                ["scope"] = scope
            });

            var tokenUrl = $"https://login.microsoftonline.com/{_options.TenantId}/oauth2/v2.0/token";

            using var response = await _http.PostAsync(tokenUrl, form, cancellationToken);
            var json = await response.Content.ReadAsStringAsync(cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                // Deliberately not logging the body - Entra error responses can
                // echo request parameters back, including the secret.
                _logger.LogError(
                    "Entra refused the {Scope} token request: {Status}. " +
                    "Check the client secret has not expired.",
                    scope, (int)response.StatusCode);

                throw new InvalidOperationException($"Bot token request failed with {(int)response.StatusCode}.");
            }

            using var doc = JsonDocument.Parse(json);
            var token = doc.RootElement.GetProperty("access_token").GetString()
                        ?? throw new InvalidOperationException("No access_token in the response.");

            if (!isGraph)
            {
                var expiresIn = doc.RootElement.TryGetProperty("expires_in", out var exp)
                    ? exp.GetInt32()
                    : 3000;

                _cachedToken = token;
                _tokenExpiresAt = DateTimeOffset.UtcNow.AddSeconds(expiresIn);
            }

            return token;
        }
        finally
        {
            _tokenLock.Release();
        }
    }

    private static string Truncate(string value, int max) =>
        value.Length <= max ? value : value[..max];
}
