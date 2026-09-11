namespace PulseNet.Approvals.Functions.Options;

public sealed class ActionTokenOptions
{
    public const string SectionName = "ActionToken";

    /// <summary>
    /// Secret used to sign action tokens. MUST be different from the dispatch
    /// signing secret: that one protects Business Central talking to us, this
    /// one protects a button in an approver's hand. Different threat, different
    /// blast radius, different key.
    /// </summary>
    public string SigningSecret { get; set; } = string.Empty;

    /// <summary>
    /// How long a button stays live. Thirty minutes is long enough for someone
    /// to finish a meeting and short enough that a screenshot in a group chat
    /// is worthless by the time it spreads.
    /// </summary>
    public int TtlMinutes { get; set; } = 30;

    /// <summary>Table Storage table holding burned nonces.</summary>
    public string NonceTable { get; set; } = "approvalactionnonces";

    /// <summary>
    /// When true the action endpoint refuses any request without an
    /// authenticated Entra principal, supplied by App Service Easy Auth.
    ///
    /// This is what turns Link mode from a sandbox toy into something
    /// production-grade: Microsoft asserts who is signed in, and we check that
    /// person is the approver the token was minted for. Someone else in a
    /// shared Teams channel tapping the button gets a clean refusal.
    /// </summary>
    public bool RequireSignedInUser { get; set; } = false;

    /// <summary>
    /// Environment tags where RequireSignedInUser may NOT be turned off.
    /// A hard guard, not a checkbox: the one configuration mistake here writes
    /// a real approval on a real invoice.
    /// </summary>
    public string EnforceSignInEnvironments { get; set; } = "PROD,PRODUCTION";
}
