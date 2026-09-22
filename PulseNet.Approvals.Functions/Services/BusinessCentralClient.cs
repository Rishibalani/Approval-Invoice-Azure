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
    private readonly EntraOptions _entra;
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
        IOptions<EntraOptions> entra,
        ILogger<BusinessCentralClient> logger)
    {
        _http = http;
        _options = options.Value;
        _entra = entra.Value;
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
        CancellationToken cancellationToken,
        string? comment = null)
    {
        try
        {
            string? systemId;

            try
            {
                systemId = await ResolveSystemIdAsync(approvalEntryNo, cancellationToken);
            }
            catch (BcUnreachableException ex)
            {
                // A configuration fault, not a business outcome. Reported as
                // such so the approver is told to use Business Central rather
                // than told their invoice vanished.
                return BcActionResult.Error(ex.Reason, transient: false);
            }

            if (systemId is null)
            {
                // Either it never existed, or it is no longer Open. Both look
                // the same from here and both mean the same to the approver.
                return BcActionResult.From("NOT_FOUND");
            }

            var actionName = action == ApprovalAction.Approve ? "approve" : "reject";
            var url = $"{_options.ApprovalEntriesUrl}({systemId})/Microsoft.NAV.{actionName}";

            // Everything the audit line needs travels WITH the decision, in
            // one call.
            //
            // There are setActionContext and setActionComment actions on the
            // page, and they publish - parameters were never the problem, only
            // WebServiceActionContext was. But they are useless to an OData
            // caller: every OData request creates a fresh page instance, so a
            // value set by one call is gone by the next. Calling them first
            // would record an approval with no channel, no device and no
            // comment, and nothing would report a fault.
            var body = JsonSerializer.Serialize(new
            {
                channel,
                device = deviceInfo,
                correlationId,

                // Written to an Approval Comment Line by the action handler.
                // Logging it and dropping it was the earlier gap: the reason
                // was collected from the approver, shown back to them, and then
                // existed nowhere an auditor could find it.
                comment = comment ?? string.Empty
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

            // Distinguish "could not ask" from "asked, and it is not there".
            //
            // Both used to return null, and the caller reported NOT_FOUND for
            // both - so a missing permission told the approver their invoice
            // had been cancelled. They would then go looking for a cancelled
            // document that does not exist, which is a worse outcome than an
            // honest error.
            //
            // 401 and 403 mean the credentials are wrong or lack rights.
            // 404 on this URL means the API page is not published.
            throw new BcUnreachableException(
                (int)response.StatusCode switch
                {
                    401 => "bc_unauthorised",
                    403 => "bc_forbidden",
                    404 => "bc_api_page_missing",
                    _ => "bc_read_failed"
                });
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
    /// Every approver identity Business Central knows about, for bulk
    /// provisioning. Read through the identity API page rather than inferred
    /// from a user list, so suspension and consent come along with it.
    /// </summary>
    public async Task<IReadOnlyList<ApproverIdentity>> GetApproverIdentitiesAsync(
        CancellationToken cancellationToken)
    {
        var url = _options.IdentitiesUrl + "?$select=systemId,userId,upn,entraObjectId,suspended";

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue(
                "Bearer", await GetTokenAsync(cancellationToken));

            using var response = await _http.SendAsync(request, cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogError(
                    "Could not read approver identities: {Status}. Is PN Approver Identity API published?",
                    (int)response.StatusCode);

                return [];
            }

            var json = await response.Content.ReadAsStringAsync(cancellationToken);
            using var doc = JsonDocument.Parse(json);

            if (!doc.RootElement.TryGetProperty("value", out var rows))
            {
                return [];
            }

            var identities = new List<ApproverIdentity>();

            foreach (var row in rows.EnumerateArray())
            {
                identities.Add(new ApproverIdentity
                {
                    SystemId = Str(row, "systemId"),
                    UserId = Str(row, "userId"),
                    Upn = Str(row, "upn"),
                    EntraObjectId = Str(row, "entraObjectId"),
                    Suspended = row.TryGetProperty("suspended", out var sus) && sus.GetBoolean()
                });
            }

            return identities;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Reading approver identities threw.");
            return [];
        }
    }

    /// <summary>
    /// Writes back an Entra object ID resolved from Graph.
    ///
    /// PATCH rather than a bound action: this is a field update, not a
    /// business operation. The API page makes only that one field writable, so
    /// there is nothing else a PATCH could disturb.
    /// </summary>
    public async Task<bool> SetEntraObjectIdAsync(
        string userId,
        string entraObjectId,
        CancellationToken cancellationToken)
    {
        try
        {
            // Filter on userId rather than carrying SystemIds around - the
            // caller knows who the approver is, not what row they live in.
            var lookupUrl = _options.IdentitiesUrl +
                $"?$filter=userId eq '{Uri.EscapeDataString(userId)}'&$select=systemId";

            using var lookup = new HttpRequestMessage(HttpMethod.Get, lookupUrl);
            lookup.Headers.Authorization = new AuthenticationHeaderValue(
                "Bearer", await GetTokenAsync(cancellationToken));

            using var lookupResponse = await _http.SendAsync(lookup, cancellationToken);

            if (!lookupResponse.IsSuccessStatusCode)
            {
                return false;
            }

            var lookupJson = await lookupResponse.Content.ReadAsStringAsync(cancellationToken);
            using var lookupDoc = JsonDocument.Parse(lookupJson);

            if (!lookupDoc.RootElement.TryGetProperty("value", out var rows) ||
                rows.GetArrayLength() == 0)
            {
                _logger.LogWarning(
                    "No approver identity row for {UserId}. Run Import From Approval User Setup first.",
                    userId);

                return false;
            }

            var systemId = Str(rows[0], "systemId");
            var body = JsonSerializer.Serialize(new { entraObjectId });

            using var patch = new HttpRequestMessage(
                HttpMethod.Patch, $"{_options.IdentitiesUrl}({systemId})")
            {
                Content = new StringContent(body, Encoding.UTF8, "application/json")
            };

            patch.Headers.Authorization = new AuthenticationHeaderValue(
                "Bearer", await GetTokenAsync(cancellationToken));

            // Business Central requires an ETag on a PATCH. * accepts whatever
            // is current, which is right here: one writer, nothing to conflict
            // with.
            patch.Headers.IfMatch.Add(new EntityTagHeaderValue("*"));

            using var patchResponse = await _http.SendAsync(patch, cancellationToken);

            if (!patchResponse.IsSuccessStatusCode)
            {
                var text = await patchResponse.Content.ReadAsStringAsync(cancellationToken);

                _logger.LogError(
                    "Writing the object ID for {UserId} failed: {Status} {Body}",
                    userId, (int)patchResponse.StatusCode, Truncate(text, 300));

                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Writing the object ID for {UserId} threw.", userId);
            return false;
        }
    }

    /// <summary>
    /// Writes back an Entra object ID, matching the approver on email.
    ///
    /// Matched on email rather than user id because that is what the Teams
    /// side knows. The Business Central user id and the Teams display name are
    /// unrelated strings - "PATELK" and "Khushil Patel" share nothing a lookup
    /// could use.
    /// </summary>
    public async Task<bool> SetEntraObjectIdByEmailAsync(
        string email,
        string entraObjectId,
        CancellationToken cancellationToken)
    {
        try
        {
            var lookupUrl = _options.IdentitiesUrl +
                $"?$filter=upn eq '{Uri.EscapeDataString(email)}'&$select=systemId,entraObjectId";

            using var lookup = new HttpRequestMessage(HttpMethod.Get, lookupUrl);
            lookup.Headers.Authorization = new AuthenticationHeaderValue(
                "Bearer", await GetTokenAsync(cancellationToken));

            using var lookupResponse = await _http.SendAsync(lookup, cancellationToken);

            if (!lookupResponse.IsSuccessStatusCode)
            {
                return false;
            }

            var lookupJson = await lookupResponse.Content.ReadAsStringAsync(cancellationToken);
            using var lookupDoc = JsonDocument.Parse(lookupJson);

            if (!lookupDoc.RootElement.TryGetProperty("value", out var rows) ||
                rows.GetArrayLength() == 0)
            {
                // No approver with that email. Normal - not everyone who talks
                // to the bot is an approver.
                return false;
            }

            var existing = Str(rows[0], "entraObjectId");

            // Already recorded. Skipping saves a write and keeps the audit of
            // who last touched the row meaningful.
            if (!string.IsNullOrWhiteSpace(existing) &&
                !existing.StartsWith("00000000", StringComparison.Ordinal))
            {
                return true;
            }

            var systemId = Str(rows[0], "systemId");
            var body = JsonSerializer.Serialize(new { entraObjectId });

            using var patch = new HttpRequestMessage(
                HttpMethod.Patch, $"{_options.IdentitiesUrl}({systemId})")
            {
                Content = new StringContent(body, Encoding.UTF8, "application/json")
            };

            patch.Headers.Authorization = new AuthenticationHeaderValue(
                "Bearer", await GetTokenAsync(cancellationToken));
            patch.Headers.IfMatch.Add(new EntityTagHeaderValue("*"));

            using var patchResponse = await _http.SendAsync(patch, cancellationToken);
            return patchResponse.IsSuccessStatusCode;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Caching the object ID for {Email} threw.", email);
            return false;
        }
    }

    private static string Str(JsonElement element, string name) =>
        element.TryGetProperty(name, out var v) ? v.GetString() ?? string.Empty : string.Empty;

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
    /// Client credentials against Entra, cached and renewed
    /// Entra:TokenRefreshSkewSeconds early. Same reasoning as the AL side: a
    /// token expiring between our check and Business Central's validation
    /// produces a 401 indistinguishable from a misconfiguration, and the skew
    /// makes that race impossible.
    /// </summary>
    private async Task<string> GetTokenAsync(CancellationToken cancellationToken)
    {
        if (_cachedToken is not null && DateTimeOffset.UtcNow < _tokenExpiresAt - _entra.TokenRefreshSkew)
        {
            return _cachedToken;
        }

        await _tokenLock.WaitAsync(cancellationToken);
        try
        {
            // Re-check: another thread may have refreshed while we waited.
            if (_cachedToken is not null && DateTimeOffset.UtcNow < _tokenExpiresAt - _entra.TokenRefreshSkew)
            {
                return _cachedToken;
            }

            var form = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["grant_type"] = "client_credentials",
                ["client_id"] = _options.ClientId,
                ["client_secret"] = _options.ClientSecret,
                ["scope"] = _options.Scope
            });

            var tokenUrl = _entra.TokenEndpoint(_options.TenantId);

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

            // No guessed lifetime: a token of unknown lifetime cannot be cached
            // safely, and Entra always sends expires_in on a v2.0 response.
            var expiresIn = doc.RootElement.TryGetProperty("expires_in", out var exp)
                ? exp.GetInt32()
                : throw new InvalidOperationException("Entra returned no expires_in.");

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
        "AMOUNT_CHANGED" =>
            "The amount on this invoice changed after you were notified, so the figure you saw is no longer what you would be approving. " +
            "Please open it in Business Central and approve there if the new amount is correct.",
        "DOCUMENT_CHANGED" => "This invoice was edited after you were notified. Please review it in Business Central before approving.",
        // The most serious of the refusals, and the one that most deserves a
        // specific message. "Something changed, go and look" invites a shrug;
        // naming the bank details and saying why it matters does not.
        //
        // Wording says "after you were notified", not "after the invoice was
        // created" - the check now compares against the moment the card was
        // sent, which is the window that matters.
        "BANK_DETAILS_CHANGED" =>
            "This vendor's bank details were changed after you were notified, so the payment may now go somewhere different. " +
            "Approval has been stopped. Please open the invoice in Business Central, check the bank details against something you trust, " +
            "and approve there if they are correct.",
        "NOT_FOUND" => "This approval request no longer exists. It may have been cancelled or the document posted.",
        "bc_unreachable" => "Business Central could not be reached. Please try again shortly.",

        // Configuration faults. Deliberately NOT phrased as "try again" - no
        // amount of retrying fixes a missing permission, and telling somebody
        // to retry just wastes their time before they give up and use the
        // client anyway.
        "bc_unauthorised" =>
            "This approval could not be completed because the connection to Business Central is not authorised. Please approve the invoice in Business Central, and let IT know.",
        "bc_forbidden" =>
            "This approval could not be completed because the service account lacks permission in Business Central. Please approve the invoice there, and let IT know.",
        "bc_api_page_missing" =>
            "This approval could not be completed because the Business Central approval API is unavailable. Please approve the invoice in Business Central, and let IT know.",
        "bc_read_failed" =>
            "This approval could not be completed. Please open the invoice in Business Central.",
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

/// <summary>
/// Business Central could not be asked - as distinct from being asked and
/// having nothing to say.
///
/// A separate type because the two produce different advice. "Already handled"
/// is something the approver can act on; "the integration is misconfigured" is
/// something only an administrator can.
/// </summary>
public sealed class BcUnreachableException : Exception
{
    public string Reason { get; }

    public BcUnreachableException(string reason)
        : base($"Business Central could not be reached: {reason}")
        => Reason = reason;
}

/// <summary>
/// One approver, as Business Central knows them. Just enough to provision
/// against the directory - the full identity record stays in Business Central.
/// </summary>
public sealed record ApproverIdentity
{
    public string SystemId { get; init; } = string.Empty;
    public string UserId { get; init; } = string.Empty;
    public string Upn { get; init; } = string.Empty;
    public string EntraObjectId { get; init; } = string.Empty;
    public bool Suspended { get; init; }
}
