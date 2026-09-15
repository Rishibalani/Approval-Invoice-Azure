using System.Security.Cryptography;
using System.Text;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Options;

namespace PulseNet.Approvals.Functions.Security;

/// <summary>
/// Verifies that a webhook really came from Meta.
///
/// The endpoint is a public URL with no key - Meta cannot attach one. Without
/// this check, anyone who found the URL could post a forged button tap and
/// approve invoices.
///
/// Meta signs the RAW request body with the app secret and sends the result as
/// X-Hub-Signature-256: sha256=<hex>.
///
/// THE USUAL WAY THIS GOES WRONG
///
/// Signing a re-serialised body. The signature covers the exact bytes Meta
/// sent, so deserialising the JSON and serialising it again changes whitespace
/// and key order, and the comparison fails for reasons invisible in a debugger.
/// Read the body once as a string and reuse that string.
/// </summary>
public sealed class WhatsAppSignatureValidator
{
    private readonly WhatsAppChannelOptions _options;
    private readonly ILogger<WhatsAppSignatureValidator> _logger;

    public WhatsAppSignatureValidator(
        IOptions<ChannelOptions> options,
        ILogger<WhatsAppSignatureValidator> logger)
    {
        _options = options.Value.WhatsApp;
        _logger = logger;
    }

    public bool IsValid(string rawBody, string? signatureHeader)
    {
        if (string.IsNullOrWhiteSpace(_options.AppSecret))
        {
            // Fail closed. A missing secret must never mean "skip the check" -
            // that turns a configuration gap into an open approval endpoint.
            _logger.LogError(
                "Channels:WhatsApp:AppSecret is not configured. Refusing all webhook traffic.");

            return false;
        }

        if (string.IsNullOrWhiteSpace(signatureHeader))
        {
            return false;
        }

        // Header format is "sha256=<hex>".
        var provided = signatureHeader.StartsWith("sha256=", StringComparison.OrdinalIgnoreCase)
            ? signatureHeader["sha256=".Length..]
            : signatureHeader;

        using var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(_options.AppSecret));
        var computed = Convert.ToHexString(hmac.ComputeHash(Encoding.UTF8.GetBytes(rawBody)))
                              .ToLowerInvariant();

        var providedBytes = Encoding.UTF8.GetBytes(provided.ToLowerInvariant());
        var computedBytes = Encoding.UTF8.GetBytes(computed);

        if (providedBytes.Length != computedBytes.Length)
        {
            return false;
        }

        // Constant time. A naive comparison returns as soon as it finds a
        // mismatched byte, and that timing leaks the correct signature one
        // byte at a time.
        var match = CryptographicOperations.FixedTimeEquals(providedBytes, computedBytes);

        if (!match)
        {
            _logger.LogWarning("WhatsApp webhook signature mismatch.");
        }

        return match;
    }

    /// <summary>
    /// The GET handshake Meta performs when you save the webhook URL. It sends
    /// a verify token and a challenge; echo the challenge back as plain text
    /// if the token matches.
    ///
    /// Returning JSON, or a quoted string, makes Meta reject the subscription
    /// with no explanation.
    /// </summary>
    public bool IsValidVerification(string? mode, string? verifyToken) =>
        mode == "subscribe" &&
        !string.IsNullOrWhiteSpace(_options.WebhookVerifyToken) &&
        verifyToken == _options.WebhookVerifyToken;
}
