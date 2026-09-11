using System.Net;
using System.Text;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Channels;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Options;
using PulseNet.Approvals.Functions.Security;
using PulseNet.Approvals.Functions.Services;

namespace PulseNet.Approvals.Functions.Functions;

/// <summary>
/// The endpoint every Link-mode button points at. One URL, every channel.
///
/// An Adaptive Card Action.OpenUrl in Teams, an anchor tag in an email, and a
/// URL button on a WhatsApp template all arrive here. That is what makes Link
/// mode one implementation rather than three.
///
/// ORDER OF CHECKS, AND WHY
///
///  1. Token signature   - before parsing anything attacker-controlled
///  2. Token expiry      - cheap, and the most common legitimate failure
///  3. Nonce not burned  - stops a replayed URL
///  4. Signed-in user    - Easy Auth, when enabled. Proves WHO is clicking
///  5. Business Central  - re-checks authority, status, amount and the gates
///  6. Burn the nonce    - only AFTER success
///
/// Step 6 is deliberately last. Burning before the Business Central call would
/// mean a transient outage permanently destroys the approver's button with no
/// way to retry. Burning after means a duplicate tap during the call is caught
/// by Business Central's own status guard, which is the right place for it.
///
/// AUTHORIZATION LEVEL IS ANONYMOUS BY DESIGN
///
/// This URL is tapped from an email client and a Teams card - neither can
/// attach a function key. The signed token IS the authentication, backed by
/// Easy Auth in production. A function key here would only stop the buttons
/// working.
/// </summary>
public sealed class ApprovalActionFunction
{
    private readonly ActionTokenService _tokenService;
    private readonly BusinessCentralClient _bcClient;
    private readonly ActionTokenOptions _tokenOptions;
    private readonly DispatchOptions _dispatchOptions;
    private readonly ILogger<ApprovalActionFunction> _logger;

    public ApprovalActionFunction(
        ActionTokenService tokenService,
        BusinessCentralClient bcClient,
        IOptions<ActionTokenOptions> tokenOptions,
        IOptions<DispatchOptions> dispatchOptions,
        ILogger<ApprovalActionFunction> logger)
    {
        _tokenService = tokenService;
        _bcClient = bcClient;
        _tokenOptions = tokenOptions.Value;
        _dispatchOptions = dispatchOptions.Value;
        _logger = logger;
    }

    [Function(nameof(ApprovalAct))]
    public async Task<IActionResult> ApprovalAct(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = "approvals/act")]
        HttpRequest req,
        CancellationToken cancellationToken)
    {
        var correlationId = Guid.NewGuid().ToString();
        var token = req.Query["t"].FirstOrDefault();

        // ---- Hard guard: never run unauthenticated in production ------
        // A configuration mistake here writes a real approval on a real
        // invoice, so this is code, not a checkbox someone can untick.
        if (RequiresSignInHere() && !_tokenOptions.RequireSignedInUser)
        {
            _logger.LogCritical(
                "Environment {Env} requires sign-in but ActionToken:RequireSignedInUser is false. Refusing.",
                _dispatchOptions.AllowedEnvironments);

            return Page(
                HttpStatusCode.InternalServerError,
                "Configuration error",
                "This service is not correctly configured for this environment. Please contact IT.",
                isError: true);
        }

        // ---- 1-3. Token ------------------------------------------------
        var validation = await _tokenService.ValidateAsync(token, cancellationToken);

        if (!validation.IsValid)
        {
            _logger.LogInformation("Action rejected: {Reason}", validation.FailureReason);

            var (title, message) = validation.FailureReason switch
            {
                "token_expired" => (
                    "This link has expired",
                    $"Approval links stay active for {_tokenOptions.TtlMinutes} minutes. Please open the invoice in Business Central to approve it."),
                "token_already_used" => (
                    "Already handled",
                    "This request has already been actioned — either by you, or by someone else in the approval chain."),
                _ => (
                    "This link is not valid",
                    "Please open the invoice in Business Central instead.")
            };

            return Page(HttpStatusCode.OK, title, message, isError: false);
        }

        // ---- 4. Who is actually clicking -------------------------------
        // Easy Auth injects this after Entra sign-in. It is what makes a card
        // posted to a shared Teams channel safe: everyone can see it, only the
        // named approver can act on it.
        var signedInUpn = req.Headers["X-MS-CLIENT-PRINCIPAL-NAME"].FirstOrDefault();

        if (_tokenOptions.RequireSignedInUser)
        {
            if (string.IsNullOrWhiteSpace(signedInUpn))
            {
                _logger.LogError(
                    "RequireSignedInUser is on but no principal header arrived. " +
                    "Is App Service Authentication actually enabled?");

                return Page(
                    HttpStatusCode.Unauthorized,
                    "Please sign in",
                    "We could not confirm who you are. Please try again from the original message.",
                    isError: true);
            }

            if (!_tokenService.MatchesSignedInUser(validation.ApproverHash, signedInUpn))
            {
                // Someone else in the channel tapped the button. Expected in a
                // shared channel, and exactly what this check is for.
                _logger.LogWarning(
                    "Approval entry {EntryNo} was tapped by a user it was not assigned to.",
                    validation.ApprovalEntryNo);

                return Page(
                    HttpStatusCode.Forbidden,
                    "Not assigned to you",
                    "This approval is assigned to someone else. If you believe that is wrong, please check in Business Central.",
                    isError: false);
            }
        }

        // ---- 5. Business Central ---------------------------------------
        var result = await _bcClient.ExecuteApprovalAsync(
            validation.ApprovalEntryNo,
            validation.Action,
            channel: DetectChannel(req),
            deviceInfo: DetectDevice(req),
            correlationId: correlationId,
            cancellationToken: cancellationToken);

        // ---- 6. Burn the nonce, only on success ------------------------
        if (result.Succeeded)
        {
            await _tokenService.BurnNonceAsync(
                validation.Nonce,
                validation.ApprovalEntryNo,
                cancellationToken);

            var verb = validation.Action == ApprovalAction.Approve ? "Approved" : "Rejected";

            _logger.LogInformation(
                "{Verb} entry {EntryNo} by {User} via {Channel}.",
                verb, validation.ApprovalEntryNo, signedInUpn ?? "(unauthenticated)", DetectChannel(req));

            return Page(HttpStatusCode.OK, verb, result.ApproverMessage, isError: false);
        }

        // A business refusal is not an error. "Somebody already approved this"
        // deserves a calm page, not a red one.
        return Page(
            result.IsBusinessRefusal ? HttpStatusCode.OK : HttpStatusCode.ServiceUnavailable,
            result.IsBusinessRefusal ? "No action taken" : "Something went wrong",
            result.ApproverMessage,
            isError: !result.IsBusinessRefusal);
    }

    // ------------------------------------------------------------------

    private bool RequiresSignInHere()
    {
        var enforced = _tokenOptions.EnforceSignInEnvironments
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        var current = _dispatchOptions.AllowedEnvironments
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        return current.Any(c => enforced.Contains(c, StringComparer.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Best-effort channel detection from the referrer, for the audit trail.
    /// Business Central writes this to the Approval Comment Line, so an auditor
    /// can see not just who approved but from where.
    /// </summary>
    private static string DetectChannel(HttpRequest req)
    {
        var referer = req.Headers["Referer"].FirstOrDefault() ?? string.Empty;

        if (referer.Contains("teams.microsoft.com", StringComparison.OrdinalIgnoreCase) ||
            referer.Contains("teams.cloud.microsoft", StringComparison.OrdinalIgnoreCase))
        {
            return "Teams";
        }

        if (referer.Contains("outlook.", StringComparison.OrdinalIgnoreCase))
        {
            return "Outlook";
        }

        return referer.Length > 0 ? "Web" : "Link";
    }

    private static string DetectDevice(HttpRequest req)
    {
        var ua = req.Headers.UserAgent.FirstOrDefault() ?? string.Empty;

        if (ua.Contains("Android", StringComparison.OrdinalIgnoreCase)) return "Android";
        if (ua.Contains("iPhone", StringComparison.OrdinalIgnoreCase)) return "iPhone";
        if (ua.Contains("iPad", StringComparison.OrdinalIgnoreCase)) return "iPad";
        if (ua.Contains("Windows", StringComparison.OrdinalIgnoreCase)) return "Windows";
        if (ua.Contains("Macintosh", StringComparison.OrdinalIgnoreCase)) return "Mac";

        return "Unknown";
    }

    /// <summary>
    /// The page the approver lands on. Deliberately plain and self-contained -
    /// it renders inside an in-app browser on a phone, where external
    /// stylesheets and fonts are unreliable.
    /// </summary>
    private static ContentResult Page(HttpStatusCode status, string title, string message, bool isError)
    {
        var accent = isError ? "#a4262c" : "#107c10";
        var safeTitle = WebUtility.HtmlEncode(title);
        var safeMessage = WebUtility.HtmlEncode(message);

        var html = new StringBuilder()
            .Append("<!DOCTYPE html><html><head><meta charset=\"utf-8\">")
            .Append("<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">")
            .Append($"<title>{safeTitle}</title></head>")
            .Append("<body style=\"margin:0;padding:0;background:#f5f5f5;")
            .Append("font-family:Segoe UI,Helvetica,Arial,sans-serif;\">")
            .Append("<div style=\"max-width:420px;margin:80px auto;background:#fff;")
            .Append("border-radius:8px;padding:32px;text-align:center;\">")
            .Append($"<div style=\"width:48px;height:4px;background:{accent};margin:0 auto 24px;border-radius:2px;\"></div>")
            .Append($"<h1 style=\"margin:0 0 12px;font-size:20px;color:#1a1a1a;\">{safeTitle}</h1>")
            .Append($"<p style=\"margin:0;font-size:15px;line-height:1.5;color:#555;\">{safeMessage}</p>")
            .Append("<p style=\"margin:28px 0 0;font-size:13px;color:#999;\">You can close this window.</p>")
            .Append("</div></body></html>")
            .ToString();

        return new ContentResult
        {
            Content = html,
            ContentType = "text/html; charset=utf-8",
            StatusCode = (int)status
        };
    }
}
