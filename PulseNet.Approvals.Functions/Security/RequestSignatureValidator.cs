using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using Azure;
using Azure.Data.Tables;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Options;

namespace PulseNet.Approvals.Functions.Security;

public sealed record SignatureValidationResult(bool IsValid, string? FailureReason)
{
    public static readonly SignatureValidationResult Success = new(true, null);
    public static SignatureValidationResult Fail(string reason) => new(false, reason);
}

/// <summary>
/// Verifies that a request really came from the paired Business Central tenant,
/// unmodified, recently, and only once.
///
/// WHY A SIGNATURE ON TOP OF A FUNCTION KEY
///
/// A function key answers "is this caller allowed in". It does not answer "is
/// this the body they sent". Anyone who obtains the key - from a log, a config
/// export, a support screenshot - can post any JSON they like, including an
/// approval for an invoice they invented. The HMAC binds the caller to the exact
/// bytes, and the bytes include the amount.
///
/// THREE CHECKS, IN THIS ORDER
///
/// 1. Clock skew. A timestamp outside the window is refused, so a captured
///    request stops being useful after a few minutes.
/// 2. Signature. Recomputed over timestamp.nonce.body using the raw body
///    string, before any deserialisation. Re-serialising a parsed object would
///    change whitespace and break the comparison - this is the single most
///    common way HMAC integrations fail.
/// 3. Nonce. Persisted and refused on second sight, which closes the replay
///    window that remains inside the skew allowance.
///
/// The comparison itself is constant-time. A naive string equality leaks the
/// correct signature one byte at a time through timing.
/// </summary>
public sealed class RequestSignatureValidator
{
    private readonly DispatchOptions _options;
    private readonly TableServiceClient _tableService;
    private readonly ILogger<RequestSignatureValidator> _logger;

    public RequestSignatureValidator(
        IOptions<DispatchOptions> options,
        TableServiceClient tableService,
        ILogger<RequestSignatureValidator> logger)
    {
        _options = options.Value;
        _tableService = tableService;
        _logger = logger;
    }

    public async Task<SignatureValidationResult> ValidateAsync(
        string rawBody,
        string? timestamp,
        string? nonce,
        string? signature,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(_options.SigningSecret))
        {
            // Fail closed. A missing secret must never mean "skip the check".
            _logger.LogError("Dispatch:SigningSecret is not configured. Refusing all requests.");
            return SignatureValidationResult.Fail("signing_not_configured");
        }

        if (string.IsNullOrWhiteSpace(timestamp) ||
            string.IsNullOrWhiteSpace(nonce) ||
            string.IsNullOrWhiteSpace(signature))
        {
            return SignatureValidationResult.Fail("missing_signature_headers");
        }

        // ---- 1. Clock skew -------------------------------------------------
        if (!TryParseTimestamp(timestamp, out var sentAt))
        {
            return SignatureValidationResult.Fail("unparseable_timestamp");
        }

        var skew = Math.Abs((DateTimeOffset.UtcNow - sentAt).TotalSeconds);
        if (skew > _options.MaxClockSkewSeconds)
        {
            _logger.LogWarning("Rejected request with clock skew of {Skew}s.", (int)skew);
            return SignatureValidationResult.Fail("timestamp_outside_window");
        }

        // ---- 2. Signature --------------------------------------------------
        var expected = ComputeSignature(timestamp, nonce, rawBody, _options.SigningSecret);

        if (!FixedTimeEquals(expected, signature))
        {
            _logger.LogWarning("Signature mismatch on inbound dispatch request.");
            return SignatureValidationResult.Fail("signature_mismatch");
        }

        // ---- 3. Replay -----------------------------------------------------
        if (await NonceAlreadySeenAsync(nonce, cancellationToken))
        {
            _logger.LogWarning("Nonce {Nonce} replayed.", nonce);
            return SignatureValidationResult.Fail("nonce_replayed");
        }

        return SignatureValidationResult.Success;
    }

    /// <summary>
    /// HMAC-SHA256 over "timestamp.nonce.body", base64 encoded.
    /// This must stay byte-identical to the AL side in
    /// PN Approval Http Client.Sign.
    /// </summary>
    public static string ComputeSignature(string timestamp, string nonce, string body, string secret)
    {
        var stringToSign = $"{timestamp}.{nonce}.{body}";

        using var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(secret));
        var hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(stringToSign));
        return Convert.ToBase64String(hash);
    }

    private static bool FixedTimeEquals(string a, string b)
    {
        var left = Encoding.UTF8.GetBytes(a);
        var right = Encoding.UTF8.GetBytes(b);

        // Length is not secret, but bailing early on it still leaks nothing
        // useful and CryptographicOperations requires equal spans.
        if (left.Length != right.Length)
        {
            return false;
        }

        return CryptographicOperations.FixedTimeEquals(left, right);
    }

    /// <summary>
    /// AL emits DateTime with Format(..., 0, 9), which is ISO 8601 round-trip.
    /// Accept a couple of shapes so a locale surprise on the BC side does not
    /// take the integration down.
    /// </summary>
    private static bool TryParseTimestamp(string value, out DateTimeOffset result)
    {
        return DateTimeOffset.TryParse(
            value,
            CultureInfo.InvariantCulture,
            DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal,
            out result);
    }

    /// <summary>
    /// Nonce store. Insert-if-absent in Table Storage: the insert itself is the
    /// check, so two concurrent replays cannot both pass.
    /// A daily partition key makes cleanup a cheap partition delete.
    /// </summary>
    private async Task<bool> NonceAlreadySeenAsync(string nonce, CancellationToken cancellationToken)
    {
        var table = _tableService.GetTableClient(_options.NonceTable);
        await table.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

        var entity = new TableEntity(DateTime.UtcNow.ToString("yyyyMMdd"), nonce)
        {
            ["SeenUtc"] = DateTime.UtcNow
        };

        try
        {
            await table.AddEntityAsync(entity, cancellationToken);
            return false;
        }
        catch (RequestFailedException ex) when (ex.Status == 409)
        {
            return true;
        }
    }
}
