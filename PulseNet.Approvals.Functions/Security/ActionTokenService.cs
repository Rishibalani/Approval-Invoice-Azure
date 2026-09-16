using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using Azure;
using Azure.Data.Tables;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Options;

namespace PulseNet.Approvals.Functions.Security;

/// <summary>
/// Mints and verifies the token that rides inside every action button.
///
/// WHAT THIS TOKEN IS FOR
///
/// The channel proves WHO tapped. This token proves WHAT THEY WERE SHOWN.
/// Both are needed. Without the token, a URL is just a URL and anyone who sees
/// it can replay it. Without the channel's identity assertion, the token alone
/// cannot tell Priya from someone reading over her shoulder.
///
/// WHY IT IS COMPACT RATHER THAN A JWT
///
/// A WhatsApp quick-reply payload is capped at 256 characters. A JWT with
/// standard claims blows past that before you add anything useful. So the
/// format is positional and terse, and the signature is truncated to 16 bytes.
///
/// Truncating an HMAC to 128 bits is fine here: forging one requires 2^128
/// work, and the token expires in thirty minutes regardless. The nonce store
/// means even a valid token only works once.
///
/// FORMAT
///
///   base64url(payload) "." base64url(hmac-sha256(payload)[0..16])
///
///   payload = 1.{entryNo}.{approverHash}.{nonce}.{expiryUnix}.{A|R}
///
/// Roughly 90 characters end to end. Comfortably inside 256, with headroom.
///
/// WHY THE APPROVER IS HASHED
///
/// The token travels in a URL, which lands in browser history, proxy logs and
/// referrer headers. Putting priya@company.com there leaks an identity on
/// every tap. A keyed hash of the UPN binds the token to her just as tightly
/// and reveals nothing - and because it is keyed, nobody can build a rainbow
/// table of your staff directory.
/// </summary>
public sealed class ActionTokenService
{
    private const string TokenVersion = "1";
    private const int SignatureBytes = 16;

    private readonly ActionTokenOptions _options;
    private readonly TableServiceClient _tableService;
    private readonly ILogger<ActionTokenService> _logger;

    public ActionTokenService(
        IOptions<ActionTokenOptions> options,
        TableServiceClient tableService,
        ILogger<ActionTokenService> logger)
    {
        _options = options.Value;
        _tableService = tableService;
        _logger = logger;
    }

    // ------------------------------------------------------------------
    //  Minting
    // ------------------------------------------------------------------

    /// <summary>
    /// Creates a token for one approval entry, one approver, one action.
    /// Approve and Reject get separate tokens with separate nonces, so burning
    /// one does not silently disable the other.
    /// </summary>
    public string Mint(int approvalEntryNo, string approverUpn, ApprovalAction action)
    {
        if (string.IsNullOrWhiteSpace(_options.SigningSecret))
        {
            throw new InvalidOperationException(
                "ActionToken:SigningSecret is not configured. Refusing to mint an unsigned token.");
        }

        var nonce = Convert.ToHexString(RandomNumberGenerator.GetBytes(8)).ToLowerInvariant();
        var expiry = DateTimeOffset.UtcNow.AddMinutes(_options.TtlMinutes).ToUnixTimeSeconds();

        var payload = string.Join('.',
            TokenVersion,
            approvalEntryNo.ToString(CultureInfo.InvariantCulture),
            HashApprover(approverUpn),
            nonce,
            expiry.ToString(CultureInfo.InvariantCulture),
            action == ApprovalAction.Approve ? "A" : "R");

        return Base64UrlEncode(Encoding.UTF8.GetBytes(payload))
             + "."
             + Base64UrlEncode(ComputeSignature(payload));
    }

    /// <summary>Full URL a Link-mode button points at.</summary>
    public string BuildActionUrl(string actionEndpointBaseUrl, string token)
    {
        var separator = actionEndpointBaseUrl.Contains('?') ? "&" : "?";
        return $"{actionEndpointBaseUrl}{separator}t={Uri.EscapeDataString(token)}";
    }

    // ------------------------------------------------------------------
    //  Verification
    // ------------------------------------------------------------------

    /// <summary>
    /// Validates a token. Does NOT burn the nonce - call BurnNonceAsync only
    /// after Business Central has accepted the action.
    ///
    /// Burning before the call would mean a transient BC failure permanently
    /// destroys the approver's button, and they would have no way to retry.
    /// Burning after means a duplicate tap during the BC call is caught by
    /// BC's own status guard instead, which is the correct place for it.
    /// </summary>
    public async Task<ActionTokenValidation> ValidateAsync(string? token, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(token))
        {
            return ActionTokenValidation.Invalid("missing_token");
        }

        var parts = token.Split('.');

        // payload has 6 dot-separated fields but is base64url encoded as one
        // unit, so the token itself is exactly two segments.
        if (parts.Length != 2)
        {
            return ActionTokenValidation.Invalid("malformed_token");
        }

        string payload;
        byte[] providedSignature;

        try
        {
            payload = Encoding.UTF8.GetString(Base64UrlDecode(parts[0]));
            providedSignature = DecodeSignature(parts[1]);
        }
        catch (FormatException)
        {
            return ActionTokenValidation.Invalid("malformed_token");
        }

        // ---- Signature first ------------------------------------------
        // Before parsing anything, because parsing attacker-controlled input
        // is exactly what you do not want to do on an unverified message.
        var expectedSignature = ComputeSignature(payload);

        if (providedSignature.Length != expectedSignature.Length ||
            !CryptographicOperations.FixedTimeEquals(providedSignature, expectedSignature))
        {
            _logger.LogWarning("Action token signature mismatch.");
            return ActionTokenValidation.Invalid("signature_mismatch");
        }

        // ---- Now it is safe to parse ----------------------------------
        var fields = payload.Split('.');
        if (fields.Length != 6 || fields[0] != TokenVersion)
        {
            return ActionTokenValidation.Invalid("unsupported_token_version");
        }

        if (!int.TryParse(fields[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var entryNo) ||
            !long.TryParse(fields[4], NumberStyles.Integer, CultureInfo.InvariantCulture, out var expiryUnix))
        {
            return ActionTokenValidation.Invalid("malformed_token");
        }

        var approverHash = fields[2];
        var nonce = fields[3];
        var action = fields[5] == "A" ? ApprovalAction.Approve : ApprovalAction.Reject;

        // ---- Expiry ----------------------------------------------------
        if (DateTimeOffset.FromUnixTimeSeconds(expiryUnix) < DateTimeOffset.UtcNow)
        {
            // Expected and routine - somebody opened an old notification.
            // Information, not a warning; this is not an attack signal.
            _logger.LogInformation("Expired action token for entry {EntryNo}.", entryNo);
            return ActionTokenValidation.Invalid("token_expired");
        }

        // ---- Replay ----------------------------------------------------
        if (await IsNonceBurnedAsync(nonce, cancellationToken))
        {
            _logger.LogInformation("Action token already used for entry {EntryNo}.", entryNo);
            return ActionTokenValidation.Invalid("token_already_used");
        }

        return ActionTokenValidation.Valid(entryNo, approverHash, nonce, action);
    }

    /// <summary>
    /// Confirms the signed-in user is the approver this token was minted for.
    ///
    /// Only meaningful when Easy Auth is on. This is the check that makes a
    /// button posted to a shared Teams channel safe: everyone can see it,
    /// only the named approver can act on it.
    /// </summary>
    public bool MatchesSignedInUser(string approverHash, string signedInUpn) =>
        !string.IsNullOrWhiteSpace(signedInUpn) &&
        string.Equals(approverHash, HashApprover(signedInUpn), StringComparison.Ordinal);

    /// <summary>
    /// Burns a nonce. The insert IS the check - Table Storage returns 409 on a
    /// duplicate row key, so two concurrent taps cannot both succeed. A
    /// read-then-write would have a race.
    /// </summary>
    public async Task<bool> BurnNonceAsync(string nonce, int entryNo, CancellationToken cancellationToken)
    {
        var table = _tableService.GetTableClient(_options.NonceTable);
        await table.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

        var entity = new TableEntity(DateTime.UtcNow.ToString("yyyyMMdd"), nonce)
        {
            ["EntryNo"] = entryNo,
            ["BurnedUtc"] = DateTime.UtcNow
        };

        try
        {
            await table.AddEntityAsync(entity, cancellationToken);
            return true;
        }
        catch (RequestFailedException ex) when (ex.Status == 409)
        {
            return false;
        }
    }

    private async Task<bool> IsNonceBurnedAsync(string nonce, CancellationToken cancellationToken)
    {
        var table = _tableService.GetTableClient(_options.NonceTable);
        await table.CreateIfNotExistsAsync(cancellationToken: cancellationToken);

        // A token lives 30 minutes, so it can only ever appear in today's or
        // yesterday's partition. Checking two partitions beats a table scan.
        foreach (var partition in new[]
                 {
                     DateTime.UtcNow.ToString("yyyyMMdd"),
                     DateTime.UtcNow.AddDays(-1).ToString("yyyyMMdd")
                 })
        {
            try
            {
                await table.GetEntityAsync<TableEntity>(partition, nonce, cancellationToken: cancellationToken);
                return true;
            }
            catch (RequestFailedException ex) when (ex.Status == 404)
            {
                // Not in this partition. Keep looking.
            }
        }

        return false;
    }

    // ------------------------------------------------------------------
    //  Primitives
    // ------------------------------------------------------------------

    private byte[] ComputeSignature(string payload)
    {
        using var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(_options.SigningSecret));
        var full = hmac.ComputeHash(Encoding.UTF8.GetBytes(payload));
        return full[..SignatureBytes];
    }

    /// <summary>
    /// Keyed hash of the approver's UPN, 8 hex characters.
    ///
    /// Keyed, not plain: a plain SHA-256 of a UPN is trivially reversible with
    /// a staff list. Keying it means an attacker holding a token still cannot
    /// work out whose it is.
    /// </summary>
    private string HashApprover(string upn)
    {
        using var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(_options.SigningSecret));
        var hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(upn.Trim().ToLowerInvariant()));
        return Convert.ToHexString(hash[..4]).ToLowerInvariant();
    }

    private static string Base64UrlEncode(byte[] input) =>
        Convert.ToBase64String(input).TrimEnd('=').Replace('+', '-').Replace('/', '_');

    /// <summary>
    /// Reads a signature written either as base64url or as hex.
    ///
    /// WHY TWO ENCODINGS
    ///
    /// This service mints base64url. Business Central mints hex, because AL
    /// produces a hex HMAC in a single call while producing base64url of a
    /// TRUNCATED HMAC would mean converting hex to bytes to base64 by hand -
    /// about thirty lines of bit-shuffling for no behavioural gain.
    ///
    /// Both are the same first sixteen bytes of the same HMAC, written
    /// differently. Accepting both here is a few lines; forcing one encoding
    /// on AL would be thirty fragile ones there.
    ///
    /// Detection is by shape rather than by a flag in the token. A truncated
    /// signature is always sixteen bytes - exactly thirty-two hex characters,
    /// or twenty-two base64url characters - so the two cannot be confused and
    /// no caller has to declare which it used.
    /// </summary>
    private static byte[] DecodeSignature(string value)
    {
        if (value.Length == SignatureBytes * 2 && IsHex(value))
        {
            return Convert.FromHexString(value);
        }

        return Base64UrlDecode(value);
    }

    private static bool IsHex(string value)
    {
        foreach (var c in value)
        {
            var isHexDigit = (c >= '0' && c <= '9')
                          || (c >= 'a' && c <= 'f')
                          || (c >= 'A' && c <= 'F');

            if (!isHexDigit)
            {
                return false;
            }
        }

        return true;
    }

    private static byte[] Base64UrlDecode(string input)
    {
        var padded = input.Replace('-', '+').Replace('_', '/');
        padded += (padded.Length % 4) switch { 2 => "==", 3 => "=", _ => string.Empty };
        return Convert.FromBase64String(padded);
    }
}

public sealed record ActionTokenValidation
{
    public required bool IsValid { get; init; }
    public string? FailureReason { get; init; }

    public int ApprovalEntryNo { get; init; }
    public string ApproverHash { get; init; } = string.Empty;
    public string Nonce { get; init; } = string.Empty;
    public ApprovalAction Action { get; init; }

    public static ActionTokenValidation Invalid(string reason) =>
        new() { IsValid = false, FailureReason = reason };

    public static ActionTokenValidation Valid(int entryNo, string approverHash, string nonce, ApprovalAction action) =>
        new()
        {
            IsValid = true,
            ApprovalEntryNo = entryNo,
            ApproverHash = approverHash,
            Nonce = nonce,
            Action = action
        };
}
