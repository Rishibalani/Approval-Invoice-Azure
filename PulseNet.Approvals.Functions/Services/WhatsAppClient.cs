using System.Globalization;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Options;

namespace PulseNet.Approvals.Functions.Services;

/// <summary>
/// Talks to the WhatsApp Cloud API.
///
/// TWO KINDS OF MESSAGE, AND THE RULE THAT SEPARATES THEM
///
/// Business-initiated messages must use a template Meta has pre-approved.
/// Free-form text is only permitted inside a 24-hour service window, which
/// opens when the recipient sends something - including tapping a quick-reply
/// button.
///
/// That rule is what shapes the rejection flow. The approval request is a
/// template because we start the conversation. The "why are you rejecting
/// this?" prompt is free text, which is only legal because their tap on Reject
/// opened the window a moment earlier.
///
/// THE 256-CHARACTER PAYLOAD
///
/// A quick-reply button carries a payload capped at 256 characters. That is
/// the constraint that made the action token compact rather than a JWT - a JWT
/// with standard claims blows past it before carrying anything useful.
///
/// Ours runs about ninety characters. Do not let it grow.
/// </summary>
public sealed class WhatsAppClient
{
    private readonly HttpClient _http;
    private readonly WhatsAppChannelOptions _options;
    private readonly ILogger<WhatsAppClient> _logger;

    public WhatsAppClient(
        HttpClient http,
        IOptions<ChannelOptions> options,
        ILogger<WhatsAppClient> logger)
    {
        _http = http;
        _options = options.Value.WhatsApp;
        _logger = logger;
    }

    /// <summary>
    /// {GraphApiBaseUrl}/{ApiVersion}/{PhoneNumberId}/messages - all three from
    /// Channels:WhatsApp configuration.
    /// </summary>
    private string SendUrl =>
        $"{_options.GraphApiBaseUrl.TrimEnd('/')}/{_options.ApiVersion}/{_options.PhoneNumberId}/messages";

    // ------------------------------------------------------------------
    //  Template send - the approval request
    // ------------------------------------------------------------------

    /// <summary>
    /// Sends the approval template.
    ///
    /// bodyParameters fill the {{1}}, {{2}}... placeholders IN ORDER. The order
    /// here must match the template exactly as approved by Meta - there are no
    /// names, only positions, so swapping two arguments silently produces a
    /// card showing the vendor where the amount should be.
    ///
    /// buttonPayloads are indexed the same way: quick-reply button 0, then 1.
    /// </summary>
    public async Task<WhatsAppSendResult> SendTemplateAsync(
        string toPhoneE164,
        IReadOnlyList<string> bodyParameters,
        IReadOnlyList<string> quickReplyPayloads,
        string? urlButtonSuffix,
        CancellationToken cancellationToken)
    {
        var components = new JsonArray
        {
            new JsonObject
            {
                ["type"] = "body",
                ["parameters"] = new JsonArray(
                    bodyParameters.Select(v => (JsonNode)new JsonObject
                    {
                        ["type"] = "text",
                        ["text"] = Sanitise(v)
                    }).ToArray())
            }
        };

        // Quick-reply buttons. The payload is what comes back on the webhook,
        // so this is where the signed action token rides.
        for (var i = 0; i < quickReplyPayloads.Count; i++)
        {
            components.Add(new JsonObject
            {
                ["type"] = "button",
                ["sub_type"] = "quick_reply",
                ["index"] = i.ToString(CultureInfo.InvariantCulture),
                ["parameters"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["type"] = "payload",
                        ["payload"] = quickReplyPayloads[i]
                    }
                }
            });
        }

        // A dynamic URL button appends a suffix to the base URL fixed in the
        // template. The base cannot be changed at send time - Meta approved it.
        if (!string.IsNullOrWhiteSpace(urlButtonSuffix))
        {
            components.Add(new JsonObject
            {
                ["type"] = "button",
                ["sub_type"] = "url",
                ["index"] = quickReplyPayloads.Count.ToString(CultureInfo.InvariantCulture),
                ["parameters"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["type"] = "text",
                        ["text"] = urlButtonSuffix
                    }
                }
            });
        }

        var body = new JsonObject
        {
            ["messaging_product"] = "whatsapp",
            ["recipient_type"] = "individual",
            ["to"] = NormalisePhone(toPhoneE164),
            ["type"] = "template",
            ["template"] = new JsonObject
            {
                ["name"] = _options.TemplateName,
                ["language"] = new JsonObject { ["code"] = _options.TemplateLanguage },
                ["components"] = components
            }
        };

        return await PostAsync(body, cancellationToken);
    }

    // ------------------------------------------------------------------
    //  Free-form text - only inside the 24-hour window
    // ------------------------------------------------------------------

    /// <summary>
    /// Sends plain text. Only legal inside the service window, which the
    /// recipient's own last message opened.
    ///
    /// Outside it Meta returns error 131047, and the caller should treat that
    /// as "the conversation has gone cold" rather than a transport fault - the
    /// remedy is a template, not a retry.
    /// </summary>
    public async Task<WhatsAppSendResult> SendTextAsync(
        string toPhoneE164,
        string text,
        CancellationToken cancellationToken)
    {
        var body = new JsonObject
        {
            ["messaging_product"] = "whatsapp",
            ["recipient_type"] = "individual",
            ["to"] = NormalisePhone(toPhoneE164),
            ["type"] = "text",
            ["text"] = new JsonObject
            {
                // Meta renders URLs in free text as links unless told not to.
                // Off, because a deep link that expands into a preview card
                // makes the conversation harder to read, not easier.
                ["preview_url"] = false,
                ["body"] = Sanitise(text)
            }
        };

        return await PostAsync(body, cancellationToken);
    }

    /// <summary>
    /// Marks an inbound message read, so the approver sees the blue ticks and
    /// knows their tap registered. Cosmetic, cheap, and its absence makes the
    /// bot feel broken during the seconds before a reply arrives.
    /// </summary>
    public async Task MarkReadAsync(string messageId, CancellationToken cancellationToken)
    {
        var body = new JsonObject
        {
            ["messaging_product"] = "whatsapp",
            ["status"] = "read",
            ["message_id"] = messageId
        };

        try
        {
            await PostAsync(body, cancellationToken);
        }
        catch (Exception ex)
        {
            // Never let a read receipt failure affect the actual work.
            _logger.LogDebug(ex, "Could not mark {MessageId} read.", messageId);
        }
    }

    // ------------------------------------------------------------------

    private async Task<WhatsAppSendResult> PostAsync(JsonObject body, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(_options.PhoneNumberId) ||
            string.IsNullOrWhiteSpace(_options.AccessToken))
        {
            return WhatsAppSendResult.Fail("whatsapp_not_configured");
        }

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, SendUrl)
            {
                Content = new StringContent(body.ToJsonString(), Encoding.UTF8, "application/json")
            };

            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _options.AccessToken);

            using var response = await _http.SendAsync(request, cancellationToken);
            var responseText = await response.Content.ReadAsStringAsync(cancellationToken);

            if (response.IsSuccessStatusCode)
            {
                using var doc = JsonDocument.Parse(responseText);

                var messageId = doc.RootElement.TryGetProperty("messages", out var messages) &&
                                messages.GetArrayLength() > 0 &&
                                messages[0].TryGetProperty("id", out var id)
                    ? id.GetString()
                    : null;

                return WhatsAppSendResult.Ok(messageId);
            }

            // Meta puts a numeric code in the body that says far more than the
            // HTTP status does. Surfacing it turns "400 Bad Request" into
            // something actionable.
            var (code, message) = ParseError(responseText);

            _logger.LogError(
                "WhatsApp send failed: HTTP {Status}, Meta code {Code}: {Message}",
                (int)response.StatusCode, code, message);

            return code switch
            {
                131047 => WhatsAppSendResult.Fail("outside_service_window"),
                131026 => WhatsAppSendResult.Fail("recipient_not_on_whatsapp"),
                132000 or 132001 or 132005 or 132007 or 132012 or 132015 =>
                    WhatsAppSendResult.Fail($"template_problem_{code}"),
                130429 or 131048 => WhatsAppSendResult.Fail("rate_limited", transient: true),
                _ => WhatsAppSendResult.Fail(
                    $"whatsapp_error_{code}",
                    transient: (int)response.StatusCode >= 500)
            };
        }
        catch (HttpRequestException ex)
        {
            _logger.LogError(ex, "Could not reach the WhatsApp Cloud API.");
            return WhatsAppSendResult.Fail("whatsapp_unreachable", transient: true);
        }
        catch (TaskCanceledException ex)
        {
            _logger.LogError(ex, "WhatsApp send timed out.");
            return WhatsAppSendResult.Fail("whatsapp_timeout", transient: true);
        }
    }

    private static (int Code, string Message) ParseError(string responseText)
    {
        try
        {
            using var doc = JsonDocument.Parse(responseText);

            if (!doc.RootElement.TryGetProperty("error", out var error))
            {
                return (0, responseText.Length > 200 ? responseText[..200] : responseText);
            }

            var code = error.TryGetProperty("code", out var c) ? c.GetInt32() : 0;
            var message = error.TryGetProperty("message", out var m) ? m.GetString() ?? "" : "";

            // error_data.details is where Meta explains a template rejection,
            // and it is considerably more useful than the top-level message.
            if (error.TryGetProperty("error_data", out var data) &&
                data.TryGetProperty("details", out var details))
            {
                message += " | " + details.GetString();
            }

            return (code, message);
        }
        catch (JsonException)
        {
            return (0, "unparseable error response");
        }
    }

    /// <summary>
    /// Meta wants digits only - no plus, no spaces, no dashes. E.164 with the
    /// leading plus stripped.
    /// </summary>
    private static string NormalisePhone(string phone) =>
        new(phone.Where(char.IsDigit).ToArray());

    /// <summary>
    /// Template body parameters may not contain newlines, tabs, or runs of
    /// more than four consecutive spaces. Meta rejects the whole send if they
    /// do, with an error that does not say which parameter was at fault.
    ///
    /// Vendor names and descriptions come from a database somebody else can
    /// write to, so this is not hypothetical.
    /// </summary>
    private static string Sanitise(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return "-";
        }

        var collapsed = new StringBuilder(value.Length);
        var spaceRun = 0;

        foreach (var ch in value)
        {
            if (ch is '\n' or '\r' or '\t')
            {
                if (spaceRun == 0)
                {
                    collapsed.Append(' ');
                    spaceRun = 1;
                }
                continue;
            }

            if (ch == ' ')
            {
                spaceRun++;
                if (spaceRun > 4) continue;
            }
            else
            {
                spaceRun = 0;
            }

            collapsed.Append(ch);
        }

        var result = collapsed.ToString().Trim();
        return result.Length == 0 ? "-" : result;
    }
}

public sealed record WhatsAppSendResult
{
    public required bool Succeeded { get; init; }
    public string? MessageId { get; init; }
    public string? FailureReason { get; init; }
    public bool IsTransient { get; init; }

    public static WhatsAppSendResult Ok(string? messageId) =>
        new() { Succeeded = true, MessageId = messageId };

    public static WhatsAppSendResult Fail(string reason, bool transient = false) =>
        new() { Succeeded = false, FailureReason = reason, IsTransient = transient };
}
