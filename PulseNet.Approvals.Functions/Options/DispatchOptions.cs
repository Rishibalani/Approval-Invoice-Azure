namespace PulseNet.Approvals.Functions.Options;

/// <summary>
/// Everything the ingest endpoint needs, bound from configuration.
///
/// All of it is runtime-changeable: in production these come from Azure App
/// Configuration with a Key Vault reference for the secret, so rotating the
/// signing secret or adding an allowed environment is a config change, not a
/// redeploy. That mirrors the Business Central side, where the same values
/// live on a setup page rather than in code.
/// </summary>
public sealed class DispatchOptions
{
    public const string SectionName = "Dispatch";

    /// <summary>
    /// Shared secret for the HMAC check. Must byte-for-byte match what was
    /// stored in Business Central Isolated Storage.
    /// </summary>
    public string SigningSecret { get; set; } = string.Empty;

    /// <summary>
    /// Comma-separated environment tags this Function will accept. A production
    /// Function lists only PROD, so a sandbox can never raise a real approval
    /// card - the single cheapest guard against a test tenant paying an invoice.
    /// </summary>
    public string AllowedEnvironments { get; set; } = string.Empty;

    /// <summary>Comma-separated BC tenant GUIDs. Empty means accept any.</summary>
    public string AllowedTenantIds { get; set; } = string.Empty;

    /// <summary>
    /// How far the x-pn-timestamp may be from now before the request is treated
    /// as a replay. Five minutes absorbs normal clock drift without giving an
    /// attacker a useful window.
    /// </summary>
    public int MaxClockSkewSeconds { get; set; } = 300;

    public string QueueName { get; set; } = "approval-dispatch";
    public string IdempotencyTable { get; set; } = "approvaldedupe";
    public string NonceTable { get; set; } = "approvalnonces";

    /// <summary>
    /// Schema versions this build understands. Anything else is rejected with a
    /// 400 rather than parsed optimistically - a silently misread money field is
    /// far worse than a refused request.
    /// </summary>
    public string SupportedSchemaVersions { get; set; } = "1.0";

    public IReadOnlySet<string> EnvironmentAllowList =>
        Split(AllowedEnvironments);

    public IReadOnlySet<string> TenantAllowList =>
        Split(AllowedTenantIds);

    public IReadOnlySet<string> SchemaAllowList =>
        Split(SupportedSchemaVersions);

    private static HashSet<string> Split(string value) =>
        value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
             .ToHashSet(StringComparer.OrdinalIgnoreCase);
}
