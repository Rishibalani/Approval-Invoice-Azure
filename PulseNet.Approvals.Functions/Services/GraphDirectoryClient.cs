using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Options;

namespace PulseNet.Approvals.Functions.Services;

/// <summary>
/// Bulk directory and Teams app operations, for provisioning approvers in one
/// pass rather than one at a time as dispatches happen.
///
/// TWO PERMISSIONS, BOTH APPLICATION, BOTH ADMIN-CONSENTED
///
///   User.Read.All                            read object IDs
///   TeamsAppInstallation.ReadWriteForUser.All  install the app for a user
///
/// The second deserves a moment's thought before it is requested. It permits
/// installing ANY Teams app for ANY user in the tenant - not just this one.
/// Reviewers notice, and the honest answer is that there is no narrower
/// permission for proactive installation. Going in with that stated is better
/// than having it discovered.
///
/// WHY SILENT INSTALL MATTERS AT ALL
///
/// Teams refuses to let a bot open a 1:1 chat with somebody who has never
/// added the app. Without installation there is no conversation, and an
/// approver falls through to email no matter how well the rest works.
///
/// Installing also fires a conversationUpdate to the bot, which is how the
/// conversation reference gets captured - so install and reference capture are
/// the same event, not two steps.
/// </summary>
public sealed class GraphDirectoryClient
{
    private readonly HttpClient _http;
    private readonly TeamsBotOptions _options;
    private readonly EntraOptions _entra;
    private readonly ILogger<GraphDirectoryClient> _logger;

    private string? _cachedToken;
    private DateTimeOffset _tokenExpiresAt = DateTimeOffset.MinValue;
    private readonly SemaphoreSlim _tokenLock = new(1, 1);

    // Resolved once per process. The catalogue ID does not change unless the
    // app is removed and re-uploaded.
    private string? _cachedCatalogueAppId;

    public GraphDirectoryClient(
        HttpClient http,
        IOptions<TeamsBotOptions> options,
        IOptions<EntraOptions> entra,
        ILogger<GraphDirectoryClient> logger)
    {
        _http = http;
        _options = options.Value;
        _entra = entra.Value;
        _logger = logger;
    }

    /// <summary>TeamsBot:GraphBaseUrl without a trailing slash.</summary>
    private string GraphBase => _options.GraphBaseUrl.TrimEnd('/');

    // ------------------------------------------------------------------
    //  Directory
    // ------------------------------------------------------------------

    /// <summary>
    /// Every enabled member of the tenant, as UPN to object ID.
    ///
    /// One paged call rather than one lookup per approver. On a tenant of a
    /// few thousand that is a handful of requests; done per-approver it is one
    /// request each, every time an object ID is missing.
    ///
    /// Guests are excluded. A guest cannot be a Business Central approver, and
    /// including them only makes the match ambiguous.
    /// </summary>
    public async Task<Dictionary<string, string>> GetDirectoryAsync(CancellationToken cancellationToken)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        var url = $"{GraphBase}/users" +
                  "?$select=id,userPrincipalName,accountEnabled,userType" +
                  "&$filter=accountEnabled eq true and userType eq 'Member'" +
                  $"&$top={_options.GraphDirectoryPageSize}";

        var token = await GetTokenAsync(cancellationToken);
        var pages = 0;

        while (!string.IsNullOrWhiteSpace(url))
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

            using var response = await _http.SendAsync(request, cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                var body = await response.Content.ReadAsStringAsync(cancellationToken);

                _logger.LogError(
                    "Graph user listing failed on page {Page}: {Status} {Body}",
                    pages, (int)response.StatusCode, Truncate(body, 400));

                // Return what was gathered rather than nothing. A partial
                // directory still provisions most approvers, and the next run
                // picks up the rest.
                break;
            }

            var json = await response.Content.ReadAsStringAsync(cancellationToken);
            using var doc = JsonDocument.Parse(json);

            if (doc.RootElement.TryGetProperty("value", out var users))
            {
                foreach (var user in users.EnumerateArray())
                {
                    var upn = user.TryGetProperty("userPrincipalName", out var u) ? u.GetString() : null;
                    var id = user.TryGetProperty("id", out var i) ? i.GetString() : null;

                    if (!string.IsNullOrWhiteSpace(upn) && !string.IsNullOrWhiteSpace(id))
                    {
                        map[upn] = id;
                    }
                }
            }

            url = doc.RootElement.TryGetProperty("@odata.nextLink", out var next)
                ? next.GetString()
                : null;

            pages++;
        }

        _logger.LogInformation("Read {Count} directory users across {Pages} page(s).", map.Count, pages);
        return map;
    }

    // ------------------------------------------------------------------
    //  Teams app installation
    // ------------------------------------------------------------------

    /// <summary>
    /// The app's ID IN THE ORG CATALOGUE, which is NOT the bot's app ID.
    ///
    /// This trips everyone. The manifest id and the bot id are the same GUID,
    /// and neither is what the install call wants - Teams assigns a separate
    /// catalogue ID when the package is uploaded. Passing the bot ID produces
    /// a 404 that reads as though the app does not exist.
    /// </summary>
    public async Task<string?> GetCatalogueAppIdAsync(string manifestId, CancellationToken cancellationToken)
    {
        if (_cachedCatalogueAppId is not null)
        {
            return _cachedCatalogueAppId;
        }

        var url = $"{GraphBase}/appCatalogs/teamsApps" +
                  $"?$filter=externalId eq '{Uri.EscapeDataString(manifestId)}'" +
                  "&$select=id,externalId,displayName,distributionMethod";

        using var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization = new AuthenticationHeaderValue(
            "Bearer", await GetTokenAsync(cancellationToken));

        using var response = await _http.SendAsync(request, cancellationToken);

        if (!response.IsSuccessStatusCode)
        {
            _logger.LogError(
                "Could not read the Teams app catalogue: {Status}", (int)response.StatusCode);
            return null;
        }

        var json = await response.Content.ReadAsStringAsync(cancellationToken);
        using var doc = JsonDocument.Parse(json);

        if (!doc.RootElement.TryGetProperty("value", out var apps) || apps.GetArrayLength() == 0)
        {
            _logger.LogError(
                "No app in the catalogue has externalId {ManifestId}. " +
                "Has the package been uploaded to the org catalogue and set to Allowed?",
                manifestId);

            return null;
        }

        _cachedCatalogueAppId = apps[0].GetProperty("id").GetString();

        _logger.LogInformation(
            "Teams catalogue app resolved: {CatalogueId}", _cachedCatalogueAppId);

        return _cachedCatalogueAppId;
    }

    /// <summary>
    /// Installs the app for one user.
    ///
    /// A 409 means it is already installed, which is a success from the
    /// caller's point of view - provisioning is meant to be safe to re-run,
    /// and treating "already done" as a failure would make every second run
    /// look broken.
    /// </summary>
    public async Task<InstallResult> InstallForUserAsync(
        string aadObjectId,
        string catalogueAppId,
        CancellationToken cancellationToken)
    {
        var body = new JsonObject
        {
            ["teamsApp@odata.bind"] = $"{GraphBase}/appCatalogs/teamsApps/{catalogueAppId}"
        };

        var url = $"{GraphBase}/users/{aadObjectId}/teamwork/installedApps";

        using var request = new HttpRequestMessage(HttpMethod.Post, url)
        {
            Content = new StringContent(body.ToJsonString(), Encoding.UTF8, "application/json")
        };

        request.Headers.Authorization = new AuthenticationHeaderValue(
            "Bearer", await GetTokenAsync(cancellationToken));

        using var response = await _http.SendAsync(request, cancellationToken);

        if (response.IsSuccessStatusCode)
        {
            return InstallResult.Installed;
        }

        if (response.StatusCode == System.Net.HttpStatusCode.Conflict)
        {
            return InstallResult.AlreadyInstalled;
        }

        var responseBody = await response.Content.ReadAsStringAsync(cancellationToken);

        // 404 on the user usually means no Teams licence - common enough in a
        // finance team that it deserves its own outcome rather than being
        // lumped in with real failures.
        if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
        {
            _logger.LogInformation(
                "Cannot install for {ObjectId} - the user has no Teams licence, or the catalogue app ID is wrong.",
                aadObjectId);

            return InstallResult.NotEligible;
        }

        if (response.StatusCode == System.Net.HttpStatusCode.TooManyRequests)
        {
            _logger.LogWarning("Graph throttled the install. Back off and retry.");
            return InstallResult.Throttled;
        }

        _logger.LogError(
            "Install failed for {ObjectId}: {Status} {Body}",
            aadObjectId, (int)response.StatusCode, Truncate(responseBody, 400));

        return InstallResult.Failed;
    }

    // ------------------------------------------------------------------

    private async Task<string> GetTokenAsync(CancellationToken cancellationToken)
    {
        if (_cachedToken is not null && DateTimeOffset.UtcNow < _tokenExpiresAt - _entra.TokenRefreshSkew)
        {
            return _cachedToken;
        }

        await _tokenLock.WaitAsync(cancellationToken);
        try
        {
            if (_cachedToken is not null && DateTimeOffset.UtcNow < _tokenExpiresAt - _entra.TokenRefreshSkew)
            {
                return _cachedToken;
            }

            var clientId = string.IsNullOrWhiteSpace(_options.GraphClientId)
                ? _options.AppId
                : _options.GraphClientId;

            var clientSecret = string.IsNullOrWhiteSpace(_options.GraphClientSecret)
                ? _options.AppPassword
                : _options.GraphClientSecret;

            var form = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["grant_type"] = "client_credentials",
                ["client_id"] = clientId,
                ["client_secret"] = clientSecret,
                ["scope"] = _options.GraphScope
            });

            var tokenUrl = _entra.TokenEndpoint(_options.TenantId);

            using var response = await _http.PostAsync(tokenUrl, form, cancellationToken);
            var json = await response.Content.ReadAsStringAsync(cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogError(
                    "Entra refused the Graph token request: {Status}. " +
                    "Check the secret has not expired and that admin consent was granted.",
                    (int)response.StatusCode);

                throw new InvalidOperationException($"Graph token request failed with {(int)response.StatusCode}.");
            }

            using var doc = JsonDocument.Parse(json);
            _cachedToken = doc.RootElement.GetProperty("access_token").GetString()
                           ?? throw new InvalidOperationException("No access_token in the response.");

            var expiresIn = doc.RootElement.TryGetProperty("expires_in", out var exp)
                ? exp.GetInt32()
                : throw new InvalidOperationException("Entra returned no expires_in.");
            _tokenExpiresAt = DateTimeOffset.UtcNow.AddSeconds(expiresIn);

            return _cachedToken;
        }
        finally
        {
            _tokenLock.Release();
        }
    }

    private static string Truncate(string value, int max) =>
        value.Length <= max ? value : value[..max];
}

public enum InstallResult
{
    Installed,
    AlreadyInstalled,

    /// <summary>No Teams licence, or the catalogue app ID is wrong.</summary>
    NotEligible,

    Throttled,
    Failed
}
