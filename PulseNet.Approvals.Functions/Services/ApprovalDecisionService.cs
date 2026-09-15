using Microsoft.Extensions.Logging;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Security;

namespace PulseNet.Approvals.Functions.Services;

/// <summary>
/// Everything that happens between "somebody tapped a button" and "Business
/// Central has recorded it", independent of which channel the tap came from.
///
/// WHY THIS IS ONE CLASS RATHER THAN CODE IN EACH ENDPOINT
///
/// Link mode, Outlook Action.Http and the Teams bot invoke all authenticate
/// differently - a browser session, an Entra token from Outlook, a Bot
/// Framework JWT. But once the caller is established, the work is identical:
/// validate our own action token, confirm the actor is the assigned approver,
/// call Business Central, and burn the nonce only if that succeeded.
///
/// Duplicating that across three endpoints is how they quietly drift apart,
/// and the one that drifts is the one that stops burning nonces.
///
/// THE ORDER MATTERS AND IS DELIBERATE
///
///   1. Token signature   - before parsing anything attacker-controlled
///   2. Expiry            - cheap, and the commonest legitimate failure
///   3. Nonce unburned    - stops a replayed request
///   4. Actor identity    - when the channel can assert one
///   5. Business Central  - re-checks authority, status, amount, gates
///   6. Burn the nonce    - only AFTER success
///
/// Step 6 last is the one people get wrong. Burning first means a transient
/// Business Central outage permanently destroys the approver's button with no
/// way to retry. Burning after means a duplicate tap during the call is caught
/// by Business Central's own status guard, which is where that check belongs.
/// </summary>
public sealed class ApprovalDecisionService
{
    private readonly ActionTokenService _tokenService;
    private readonly BusinessCentralClient _bcClient;
    private readonly ILogger<ApprovalDecisionService> _logger;

    public ApprovalDecisionService(
        ActionTokenService tokenService,
        BusinessCentralClient bcClient,
        ILogger<ApprovalDecisionService> logger)
    {
        _tokenService = tokenService;
        _bcClient = bcClient;
        _logger = logger;
    }

    /// <summary>
    /// Validates a token and reports what it asks for, WITHOUT consuming the
    /// nonce or touching Business Central.
    ///
    /// Exists for WhatsApp. A rejection there is two messages - the tap, then
    /// the reason - and the tap must not burn the nonce, or the approver
    /// cannot finish what they started. Peek first, execute when the reason
    /// arrives.
    ///
    /// Every other channel collects the comment in the same interaction and
    /// goes straight to ExecuteAsync.
    /// </summary>
    public async Task<ActionTokenPeek> PeekAsync(string? token, CancellationToken cancellationToken)
    {
        var validation = await _tokenService.ValidateAsync(token, cancellationToken);

        if (!validation.IsValid)
        {
            return new ActionTokenPeek
            {
                IsValid = false,
                Message = validation.FailureReason switch
                {
                    "token_expired" =>
                        "This approval request has expired. Please open the invoice in Business Central.",
                    "token_already_used" =>
                        "This request has already been handled - either by you, or by someone else in the approval chain.",
                    _ =>
                        "This button is no longer valid. Please open the invoice in Business Central."
                }
            };
        }

        return new ActionTokenPeek
        {
            IsValid = true,
            ApprovalEntryNo = validation.ApprovalEntryNo,
            Action = validation.Action,
            Message = string.Empty
        };
    }

    /// <summary>
    /// Executes a decision carried by an action token.
    /// </summary>
    /// <param name="token">The signed action token from the button.</param>
    /// <param name="assertedIdentity">
    /// The identity the channel vouches for - a signed-in UPN from Easy Auth,
    /// the mailbox from an Outlook token. Null when the channel cannot assert
    /// one, in which case only the token binds the action to a person.
    /// </param>
    /// <param name="requireAssertedIdentity">
    /// When true, a missing or mismatched identity is refused rather than
    /// waved through. This is what makes a card in a shared surface safe.
    /// </param>
    public async Task<ApprovalDecisionOutcome> ExecuteAsync(
        string? token,
        string? assertedIdentity,
        bool requireAssertedIdentity,
        string channel,
        string deviceInfo,
        CancellationToken cancellationToken,
        string? comment = null,
        bool requireRejectionReason = false)
    {
        // ---- 1-3. Our own token ---------------------------------------
        var validation = await _tokenService.ValidateAsync(token, cancellationToken);

        if (!validation.IsValid)
        {
            _logger.LogInformation(
                "Action rejected on {Channel}: {Reason}", channel, validation.FailureReason);

            return validation.FailureReason switch
            {
                "token_expired" => ApprovalDecisionOutcome.Refused(
                    "This link has expired",
                    "Approval links stay active for a limited time. Please open the invoice in Business Central."),

                "token_already_used" => ApprovalDecisionOutcome.Refused(
                    "Already handled",
                    "This request has already been actioned - either by you, or by someone else in the approval chain."),

                _ => ApprovalDecisionOutcome.Refused(
                    "This link is not valid",
                    "Please open the invoice in Business Central instead.")
            };
        }

        // ---- 4. Who is actually acting --------------------------------
        if (requireAssertedIdentity)
        {
            if (string.IsNullOrWhiteSpace(assertedIdentity))
            {
                _logger.LogError(
                    "{Channel} requires an asserted identity but none arrived. " +
                    "Check that the channel's authentication is actually enabled.",
                    channel);

                return ApprovalDecisionOutcome.Refused(
                    "Please sign in",
                    "We could not confirm who you are. Please try again from the original message.");
            }

            if (!_tokenService.MatchesSignedInUser(validation.ApproverHash, assertedIdentity))
            {
                // Somebody other than the assigned approver acted on the card.
                // Expected on a shared surface, and exactly what this check is
                // here to catch.
                _logger.LogWarning(
                    "Approval entry {EntryNo} was actioned by a user it is not assigned to, on {Channel}.",
                    validation.ApprovalEntryNo, channel);

                return ApprovalDecisionOutcome.Refused(
                    "Not assigned to you",
                    "This approval is assigned to someone else. If you believe that is wrong, please check in Business Central.");
            }
        }

        // ---- 4b. Rejection reason ------------------------------------
        //
        // Enforced HERE, not on the card. Adaptive Card 1.0 accepts isRequired
        // and ignores it, so the card cannot stop an empty submission - and a
        // client-side check would not stop a crafted request in any case. The
        // server is the only place this can actually hold.
        if (requireRejectionReason &&
            validation.Action == ApprovalAction.Reject &&
            string.IsNullOrWhiteSpace(comment))
        {
            _logger.LogInformation(
                "Rejection of entry {EntryNo} refused on {Channel}: no reason given.",
                validation.ApprovalEntryNo, channel);

            return ApprovalDecisionOutcome.Refused(
                "A reason is required",
                "Please say why you are rejecting this invoice, then confirm again.");
        }

        // ---- 5. Business Central --------------------------------------
        var result = await _bcClient.ExecuteApprovalAsync(
            validation.ApprovalEntryNo,
            validation.Action,
            channel,
            deviceInfo,
            correlationId: Guid.NewGuid().ToString(),
            cancellationToken: cancellationToken,
            comment: comment);

        // ---- 6. Burn, only on success ---------------------------------
        if (result.Succeeded)
        {
            await _tokenService.BurnNonceAsync(
                validation.Nonce, validation.ApprovalEntryNo, cancellationToken);

            var verb = validation.Action == ApprovalAction.Approve ? "Approved" : "Rejected";

            _logger.LogInformation(
                "{Verb} entry {EntryNo} by {Actor} via {Channel}.",
                verb, validation.ApprovalEntryNo, assertedIdentity ?? "(unasserted)", channel);

            return ApprovalDecisionOutcome.Succeeded(
                verb,
                result.ApproverMessage,
                validation.ApprovalEntryNo,
                validation.Action);
        }

        // A business refusal is not an error. "Somebody already approved this"
        // deserves a calm message, not a red one.
        return new ApprovalDecisionOutcome
        {
            Success = false,
            IsBusinessRefusal = result.IsBusinessRefusal,
            Title = result.IsBusinessRefusal ? "No action taken" : "Something went wrong",
            Message = result.ApproverMessage,
            ApprovalEntryNo = validation.ApprovalEntryNo,
            Action = validation.Action
        };
    }
}

public sealed record ApprovalDecisionOutcome
{
    public required bool Success { get; init; }
    public required string Title { get; init; }
    public required string Message { get; init; }

    /// <summary>
    /// True when the outcome was a legitimate business refusal rather than a
    /// technical failure. Channels render these calmly - an already-handled
    /// approval is a normal event, not a fault.
    /// </summary>
    public bool IsBusinessRefusal { get; init; }

    public int ApprovalEntryNo { get; init; }
    public ApprovalAction Action { get; init; }

    public static ApprovalDecisionOutcome Refused(string title, string message) =>
        new() { Success = false, IsBusinessRefusal = true, Title = title, Message = message };

    public static ApprovalDecisionOutcome Succeeded(
        string title, string message, int entryNo, ApprovalAction action) =>
        new()
        {
            Success = true,
            Title = title,
            Message = message,
            ApprovalEntryNo = entryNo,
            Action = action
        };
}

/// <summary>
/// What a token asks for, established without committing to it.
/// </summary>
public sealed record ActionTokenPeek
{
    public required bool IsValid { get; init; }
    public required string Message { get; init; }
    public int ApprovalEntryNo { get; init; }
    public ApprovalAction Action { get; init; }

    /// <summary>
    /// Not carried by the token - it holds only the entry number, to stay
    /// inside WhatsApp's 256-character payload. Left null unless a caller
    /// looks it up.
    /// </summary>
    public string? DocumentNo { get; init; }
}
