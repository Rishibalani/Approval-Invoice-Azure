using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Options;

namespace PulseNet.Approvals.Functions.Services;

/// <summary>
/// Sends through Microsoft Graph sendMail as a dedicated mailbox.
///
/// A NOTE ON THE PERMISSION THIS NEEDS
///
/// Mail.Send as an APPLICATION permission grants the ability to send as any
/// mailbox in the tenant. Not the approvals mailbox - every mailbox, finance
/// and executives included.
///
/// Narrow it with an Exchange application access policy scoped to a group
/// containing only the sending mailbox:
///
///   New-ApplicationAccessPolicy -AppId {clientId}
///     -PolicyScopeGroupId invoice-approval-senders@yourdomain.com
///     -AccessRight RestrictAccess
///
/// Then prove it with Test-ApplicationAccessPolicy against a DIFFERENT mailbox
/// and confirm it returns Denied. Testing only the allowed mailbox tells you
/// nothing about whether the restriction actually applies.
///
/// Without that policy this class can impersonate anyone by email, and one
/// leaked secret becomes a considerably larger problem.
/// </summary>
public sealed class GraphEmailTransport : IEmailTransport
{
    private readonly HttpClient _http;
    private readonly ChannelOptions _options;
    private readonly ILogger<GraphEmailTransport> _logger;

    private string? _cachedToken;
    private DateTimeOffset _tokenExpiresAt = DateTimeOffset.MinValue;
    private readonly SemaphoreSlim _tokenLock = new(1, 1);

    public string Name => "Graph";

    public GraphEmailTransport(
        HttpClient http,
        IOptions<ChannelOptions> options,
        ILogger<GraphEmailTransport> logger)
    {
        _http = http;
        _options = options.Value;
        _logger = logger;
    }

    public async Task<EmailSendResult> SendAsync(EmailMessage message, CancellationToken cancellationToken)
    {
        var outlook = _options.Outlook;

        if (string.IsNullOrWhiteSpace(outlook.FromAddress))
        {
            return EmailSendResult.Fail("no_from_address_configured");
        }

        try
        {
            var body = new JsonObject
            {
                ["message"] = new JsonObject
                {
                    ["subject"] = message.Subject,
                    ["body"] = new JsonObject
                    {
                        ["contentType"] = "HTML",
                        ["content"] = message.HtmlBody
                    },
                    ["toRecipients"] = new JsonArray
                    {
                        new JsonObject
                        {
                            ["emailAddress"] = new JsonObject
                            {
                                ["address"] = message.ToAddress
                            }
                        }
                    }
                },

                // Keeps a copy in the sending mailbox. Worth the storage: it
                // is the record that proves a notification went out, and it is
                // the first thing anyone investigating "I never got it" asks
                // for.
                ["saveToSentItems"] = true
            };

            // sendMail as the configured mailbox, not as the signed-in user -
            // there is no signed-in user here.
            var url = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(outlook.FromAddress)}/sendMail";

            using var request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(body.ToJsonString(), Encoding.UTF8, "application/json")
            };

            request.Headers.Authorization = new AuthenticationHeaderValue(
                "Bearer", await GetTokenAsync(cancellationToken));

            using var response = await _http.SendAsync(request, cancellationToken);

            if (response.IsSuccessStatusCode)
            {
                // sendMail returns 202 with no body, so there is no message ID
                // to correlate on. Finding it later means searching Sent Items.
                return EmailSendResult.Ok();
            }

            var responseText = await response.Content.ReadAsStringAsync(cancellationToken);

            // 403 here almost always means the application access policy is
            // blocking this mailbox rather than a missing permission. Worth
            // saying so - the two look identical otherwise.
            if (response.StatusCode == System.Net.HttpStatusCode.Forbidden)
            {
                _logger.LogError(
                    "Graph refused sendMail as {From}. Check the Exchange application access policy " +
                    "includes this mailbox, and that admin consent was granted for Mail.Send. {Body}",
                    outlook.FromAddress, Truncate(responseText, 400));
            }
            else
            {
                _logger.LogError(
                    "Graph sendMail failed: {Status} {Body}",
                    (int)response.StatusCode, Truncate(responseText, 400));
            }

            return EmailSendResult.Fail(
                $"graph_http_{(int)response.StatusCode}",
                transient: (int)response.StatusCode >= 500 ||
                           response.StatusCode == System.Net.HttpStatusCode.TooManyRequests);
        }
        catch (HttpRequestException ex)
        {
            _logger.LogError(ex, "Could not reach Microsoft Graph.");
            return EmailSendResult.Fail("graph_unreachable", transient: true);
        }
        catch (TaskCanceledException ex)
        {
            _logger.LogError(ex, "Graph sendMail timed out.");
            return EmailSendResult.Fail("graph_timeout", transient: true);
        }
        catch (InvalidOperationException ex)
        {
            _logger.LogError(ex, "Graph credentials are not usable.");
            return EmailSendResult.Fail("graph_not_configured");
        }
    }

    private async Task<string> GetTokenAsync(CancellationToken cancellationToken)
    {
        if (_cachedToken is not null && DateTimeOffset.UtcNow < _tokenExpiresAt.AddMinutes(-5))
        {
            return _cachedToken;
        }

        await _tokenLock.WaitAsync(cancellationToken);
        try
        {
            if (_cachedToken is not null && DateTimeOffset.UtcNow < _tokenExpiresAt.AddMinutes(-5))
            {
                return _cachedToken;
            }

            var outlook = _options.Outlook;

            if (string.IsNullOrWhiteSpace(outlook.ClientId) ||
                string.IsNullOrWhiteSpace(outlook.ClientSecret) ||
                string.IsNullOrWhiteSpace(outlook.TenantId))
            {
                throw new InvalidOperationException(
                    "Channels:Outlook TenantId, ClientId and ClientSecret must all be set for the Graph transport.");
            }

            var form = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["grant_type"] = "client_credentials",
                ["client_id"] = outlook.ClientId,
                ["client_secret"] = outlook.ClientSecret,
                ["scope"] = "https://graph.microsoft.com/.default"
            });

            var tokenUrl = $"https://login.microsoftonline.com/{outlook.TenantId}/oauth2/v2.0/token";

            using var response = await _http.PostAsync(tokenUrl, form, cancellationToken);
            var json = await response.Content.ReadAsStringAsync(cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                // Not logging the body - Entra error responses can echo request
                // parameters back, including the secret.
                _logger.LogError(
                    "Entra refused the Graph token request: {Status}. " +
                    "Check the client secret has not expired.",
                    (int)response.StatusCode);

                throw new InvalidOperationException($"Graph token request failed with {(int)response.StatusCode}.");
            }

            using var doc = JsonDocument.Parse(json);
            var token = doc.RootElement.GetProperty("access_token").GetString()
                        ?? throw new InvalidOperationException("No access_token in the response.");

            var expiresIn = doc.RootElement.TryGetProperty("expires_in", out var exp) ? exp.GetInt32() : 3000;

            _cachedToken = token;
            _tokenExpiresAt = DateTimeOffset.UtcNow.AddSeconds(expiresIn);

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
