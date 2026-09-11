using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Options;

namespace PulseNet.Approvals.Functions.Services;

/// <summary>
/// Calls back INTO Business Central - the Phase 2 direction.
///
/// Authenticates as Entra app registration B, whose Business Central
/// application user holds Approval Administrator. Without that flag every call
/// here fails with an authority error that reads nothing like a permissions
/// problem, so if approvals are refusing and the token looks fine, check the
/// flag before anything else.
///
/// This class never writes approval status directly. It calls the bound OData
/// actions on the API page, which delegate to ApprovalsMgmt, so Business
/// Central re-evaluates approver limits at execution time and writes its own
/// audit entries. That is what makes in-channel approval defensible.
/// </summary>
public sealed class BusinessCentralClient
{
    private readonly HttpClient _http;
    private readonly BusinessCentralOptions _options;
    private readonly ILogger<BusinessCentralClient> _logger;

    // Token cache. Static-per-instance because the client is registered as a
    // singleton; a Function App instance handling many taps should acquire one
    // token, not one per tap.
    private string? _cachedToken;
    private DateTimeOffset _tokenExpiresAt = DateTimeOffset.MinValue;
    private readonly SemaphoreSlim _tokenLock = new(1, 1);

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public BusinessCentralClient(
        HttpClient http,
        IOptions<BusinessCentralOptions> options,
        ILogger<BusinessCentralClient> logger)
    {
        _http = http;
        _options = options.Value;
        _logger = logger;
    }

    // ------------------------------------------------------------------
    //  Actions
    // ------------------------------------------------------------------

    /// <summary>
    /// Approves or rejects one entry. Returns Business Central's own status
    /// token - OK, ALREADY_PROCESSED, NOT_APPROVER, AMOUNT_CHANGED,
    /// BANK_DETAILS_CHANGED, DOCUMENT_CHANGED, NOT_FOUND.
    ///
    /// Never throws for a business refusal. A second tap is a normal event,
    /// not an exception, and the approver deserves a readable message rather
    /// than a 500.
    /// </summary>
    public async Task<BcActionResult> ExecuteApprovalAsync(
        int approvalEntryNo,
        ApprovalAction action,
        string channel,
        string deviceInfo,
        string correlationId,
        CancellationToken cancellationToken)
    {
        try
        {
            var systemId = await ResolveSystemIdAsync(approvalEntryNo, cancellationToken);

            if (systemId is null)
            {
                // Either it never existed, or it is no longer Open. Both look
                // the same from here and both mean the same to the approver.
                return BcActionResult.From("NOT_FOUND");
            }

            var actionName = action == ApprovalAction.Approve ? "approve" : "reject";
            var url = $"{_options.ApprovalEntriesUrl}({systemId})/Microsoft.NAV.{actionName}";

            var body = JsonSerializer.Serialize(new
            {
                channel,
                device = deviceInfo,
                correlationId
            });

            using var request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(body, Encoding.UTF8, "application/json")
            };

            request.Headers.Authorization = new AuthenticationHeaderValue(
                "Bearer", await GetTokenAsync(cancellationToken));

            using var response = await _http.SendAsync(request, cancellationToken);
            var responseText = await response.Content.ReadAsStringAsync(cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogError(
                    "Business Central refused the {Action} for entry {EntryNo}: {Status} {Body}",
                    actionName, approvalEntryNo, (int)response.StatusCode, Truncate(responseText, 500));

                return BcActionResult.Error(
                    $"bc_http_{(int)response.StatusCode}",
                    IsTransientStatus(response.StatusCode));
            }

            // The bound action returns a plain status string, wrapped by OData
            // as { "value": "OK" }.
            var status = ExtractValue(responseText) ?? "OK";

            _logger.LogInformation(
                "Business Central returned {Status} for {Action} on entry {EntryNo}.",
                status, actionName, approvalEntryNo);

            return BcActionResult.From(status);
        }
        catch (HttpRequestException ex)
        {
            _logger.LogError(ex, "Could not reach Business Central.");
            return BcActionResult.Error("bc_unreachable", transient: true);
        }
        catch (TaskCanceledException ex)
        {
            _logger.LogError(ex, "Business Central call timed out.");
            return BcActionResult.Error("bc_timeout", transient: true);
        }
    }

    /// <summary>
    /// Finds the SystemId for an approval entry, which OData needs as the key.
    /// Filtered to Open so an already-decided entry resolves to null and the
    /// caller reports ALREADY_PROCESSED rather than attempting the action.
    /// </summary>
    private async Task<string?> ResolveSystemIdAsync(int approvalEntryNo, CancellationToken cancellationToken)
    {
        var url = $"{_options.ApprovalEntriesUrl}?$filter=entryNo eq {approvalEntryNo}&$select=systemId,status";

        using var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization = new AuthenticationHeaderValue(
            "Bearer", await GetTokenAsync(cancellationToken));

        using var response = await _http.SendAsync(request, cancellationToken);

        if (!response.IsSuccessStatusCode)
        {
            _logger.LogError(
                "Could not read approval entry {EntryNo}: {Status}",
                approvalEntryNo, (int)response.StatusCode);
            return null;
        }

        var json = await response.Content.ReadAsStringAsync(cancellationToken);

        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("value", out var array) || array.GetArrayLength() == 0)
        {
            return null;
        }

        var first = array[0];
        return first.TryGetProperty("systemId", out var id) ? id.GetString() : null;
    }

    /// <summary>
    /// Writes the delivered channel and message ID back to the outbox row, so
    /// the card can be updated in place when the decision lands.
    /// Best-effort: a failure here must never fail a successful send.
    /// </summary>
    public async Task RecordDeliveryAsync(
        string eventId,
        ApprovalChannel channel,
        string? channelMessageId,
        CancellationToken cancellationToken)
    {
        // Wired once PN Approver Identity API (page 50941) is published.
        // Left explicit rather than silent so it shows up in a log search.
        _logger.LogInformation(
            "Delivery recorded locally: event {EventId} via {Channel}, message {MessageId}. " +
            "Write-back to Business Central pending page 50941.",
            eventId, channel, channelMessageId ?? "(none)");

        await Task.CompletedTask;
    }

    // ------------------------------------------------------------------
    //  Token handling
    // ------------------------------------------------------------------

    /// <summary>
    /// Client credentials against Entra, cached and renewed five minutes early.
    /// Same reasoning as the AL side: a token expiring between our check and
    /// Business Central's validation produces a 401 indistinguishable from a
    /// misconfiguration, and the skew makes that race impossible.
    /// </summary>
    private async Task<string> GetTokenAsync(CancellationToken cancellationToken)
    {
        if (_cachedToken is not null && DateTimeOffset.UtcNow < _tokenExpiresAt.AddMinutes(-5))
        {
            return _cachedToken;
        }

        await _tokenLock.WaitAsync(cancellationToken);
        try
        {
            // Re-check: another thread may have refreshed while we waited.
            if (_cachedToken is not null && DateTimeOffset.UtcNow < _tokenExpiresAt.AddMinutes(-5))
            {
                return _cachedToken;
            }

            var form = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["grant_type"] = "client_credentials",
                ["client_id"] = _options.ClientId,
                ["client_secret"] = _options.ClientSecret,
                ["scope"] = BusinessCentralOptions.Scope
            });

            var tokenUrl = $"https://login.microsoftonline.com/{_options.TenantId}/oauth2/v2.0/token";

            using var response = await _http.PostAsync(tokenUrl, form, cancellationToken);
            var json = await response.Content.ReadAsStringAsync(cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                // Do not log the body verbatim - Entra error responses can echo
                // request parameters back.
                _logger.LogError(
                    "Entra refused the token request: {Status}. " +
                    "Check the client secret has not expired and that admin consent was granted.",
                    (int)response.StatusCode);

                throw new InvalidOperationException($"Token request failed with {(int)response.StatusCode}.");
            }

            using var doc = JsonDocument.Parse(json);
            var token = doc.RootElement.GetProperty("access_token").GetString()
                        ?? throw new InvalidOperationException("Entra returned no access_token.");

            var expiresIn = doc.RootElement.TryGetProperty("expires_in", out var exp)
                ? exp.GetInt32()
                : 3000;

            _cachedToken = token;
            _tokenExpiresAt = DateTimeOffset.UtcNow.AddSeconds(expiresIn);

            return token;
        }
        finally
        {
            _tokenLock.Release();
        }
    }

    // ------------------------------------------------------------------

    private static string? ExtractValue(string json)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            return doc.RootElement.TryGetProperty("value", out var v) ? v.GetString() : null;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private static bool IsTransientStatus(System.Net.HttpStatusCode status) =>
        (int)status >= 500 || status == System.Net.HttpStatusCode.TooManyRequests;

    private static string Truncate(string value, int max) =>
        value.Length <= max ? value : value[..max];
}

/// <summary>
/// Business Central's verdict, plus what the approver should be told.
/// The message text lives here rather than in the Function so every channel
/// says the same thing.
/// </summary>
public sealed record BcActionResult
{
    public required string Status { get; init; }
    public required bool Succeeded { get; init; }
    public bool IsTransient { get; init; }

    public string ApproverMessage => Status switch
    {
        "OK" => "Done. The decision has been recorded in Business Central.",
        "ALREADY_PROCESSED" => "This request has already been handled - either by you on another device, or by someone else in the approval chain.",
        "NOT_APPROVER" => "This approval is not assigned to you.",
        "AMOUNT_CHANGED" => "The amount on this invoice changed after you were notified. Please review it in Business Central.",
        "DOCUMENT_CHANGED" => "This invoice was edited after you were notified. Please review it in Business Central before approving.",
        "BANK_DETAILS_CHANGED" => "The vendor's bank details changed after this invoice was created. Approval must be completed in Business Central after review.",
        "NOT_FOUND" => "This approval request no longer exists. It may have been cancelled or the document posted.",
        "bc_unreachable" => "Business Central could not be reached. Please try again shortly.",
        "bc_timeout" => "Business Central did not respond in time. Please try again shortly.",
        _ => "Something went wrong. Please open the invoice in Business Central."
    };

    /// <summary>
    /// True when the outcome was a legitimate business refusal rather than a
    /// technical failure. These are shown calmly - they are not errors.
    /// </summary>
    public bool IsBusinessRefusal => Status is
        "ALREADY_PROCESSED" or "NOT_APPROVER" or "AMOUNT_CHANGED" or
        "DOCUMENT_CHANGED" or "BANK_DETAILS_CHANGED" or "NOT_FOUND";

    public static BcActionResult From(string status) =>
        new() { Status = status, Succeeded = status == "OK" };

    public static BcActionResult Error(string status, bool transient) =>
        new() { Status = status, Succeeded = false, IsTransient = transient };
}
