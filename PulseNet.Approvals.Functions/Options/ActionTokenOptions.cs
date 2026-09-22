using System.ComponentModel.DataAnnotations;

namespace PulseNet.Approvals.Functions.Options;

/// <summary>
/// Action token settings. Every value is bound from the ActionToken section
/// and validated at startup - there are no code-level defaults. See
/// Options/Validation/ActionTokenOptionsValidator for the checks that
/// DataAnnotations cannot express.
/// </summary>
public sealed class ActionTokenOptions
{
    public const string SectionName = "ActionToken";

    /// <summary>
    /// Secret used to sign action tokens. MUST be different from the dispatch
    /// signing secret: that one protects Business Central talking to us, this
    /// one protects a button in an approver's hand. Different threat, different
    /// blast radius, different key.
    /// </summary>
    [Required(ErrorMessage = "ActionToken__SigningSecret is required.")]
    public string SigningSecret { get; set; } = string.Empty;

    /// <summary>
    /// How long a button stays live, in minutes. Short enough that a
    /// screenshot in a group chat is worthless by the time it spreads.
    ///
    /// Capped at 1440 (one day) because the burned-nonce lookup only checks
    /// today's and yesterday's daily partitions - a longer TTL could let a
    /// token outlive the partition its nonce was burned into.
    /// </summary>
    [Range(1, 1440, ErrorMessage = "ActionToken__TtlMinutes must be > 0 and <= 1440.")]
    public int TtlMinutes { get; set; }

    /// <summary>Table Storage table holding burned nonces.</summary>
    [Required(ErrorMessage = "ActionToken__NonceTable is required.")]
    public string NonceTable { get; set; } = string.Empty;

    /// <summary>
    /// When true the action endpoint refuses any request without an
    /// authenticated Entra principal, supplied by App Service Easy Auth.
    ///
    /// This is what turns Link mode from a sandbox toy into something
    /// production-grade: Microsoft asserts who is signed in, and we check that
    /// person is the approver the token was minted for. Someone else in a
    /// shared Teams channel tapping the button gets a clean refusal.
    ///
    /// Must be present in configuration (validated at startup) - a missing
    /// key must never silently mean "off".
    /// </summary>
    public bool RequireSignedInUser { get; set; }

    /// <summary>
    /// Comma-separated environment tags where RequireSignedInUser may NOT be
    /// turned off. A hard guard, not a checkbox: the one configuration mistake
    /// here writes a real approval on a real invoice.
    /// </summary>
    [Required(ErrorMessage = "ActionToken__EnforceSignInEnvironments is required.")]
    public string EnforceSignInEnvironments { get; set; } = string.Empty;
}
