using System.ComponentModel.DataAnnotations;

namespace PulseNet.Approvals.Functions.Options;

/// <summary>
/// Everything the ingest endpoint needs, bound from configuration.
///
/// All of it is runtime-changeable: in production these come from Azure App
/// Configuration with a Key Vault reference for the secret, so rotating the
/// signing secret or adding an allowed environment is a config change, not a
/// redeploy. That mirrors the Business Central side, where the same values
/// live on a setup page rather than in code.
///
/// There are no code-level defaults: every required value is validated at
/// startup and the app refuses to start without it.
/// </summary>
public sealed class DispatchOptions
{
    public const string SectionName = "Dispatch";

    /// <summary>
    /// Shared secret for the HMAC check. Must byte-for-byte match what was
    /// stored in Business Central Isolated Storage.
    /// </summary>
    [Required(ErrorMessage = "Dispatch__SigningSecret is required.")]
    public string SigningSecret { get; set; } = string.Empty;

    /// <summary>
    /// Comma-separated environment tags this Function will accept. A production
    /// Function lists only PROD, so a sandbox can never raise a real approval
    /// card - the single cheapest guard against a test tenant paying an invoice.
    ///
    /// Required: a missing allow-list must not silently mean "accept any".
    /// </summary>
    [Required(ErrorMessage = "Dispatch__AllowedEnvironments is required.")]
    public string AllowedEnvironments { get; set; } = string.Empty;

    /// <summary>Comma-separated BC tenant GUIDs. Empty means accept any.</summary>
    public string AllowedTenantIds { get; set; } = string.Empty;

    /// <summary>
    /// How far the x-pn-timestamp may be from now before the request is treated
    /// as a replay. A few minutes absorbs normal clock drift without giving an
    /// attacker a useful window.
    /// </summary>
    [Range(1, int.MaxValue, ErrorMessage = "Dispatch__MaxClockSkewSeconds must be > 0.")]
    public int MaxClockSkewSeconds { get; set; }

    /// <summary>
    /// Storage queue between ingest and the dispatch worker. Also read by the
    /// queue trigger as %Dispatch:QueueName%, so both sides use one setting.
    /// </summary>
    [Required(ErrorMessage = "Dispatch__QueueName is required.")]
    public string QueueName { get; set; } = string.Empty;

    [Required(ErrorMessage = "Dispatch__IdempotencyTable is required.")]
    public string IdempotencyTable { get; set; } = string.Empty;

    [Required(ErrorMessage = "Dispatch__NonceTable is required.")]
    public string NonceTable { get; set; } = string.Empty;

    /// <summary>
    /// Schema versions this build understands. Anything else is rejected with a
    /// 400 rather than parsed optimistically - a silently misread money field is
    /// far worse than a refused request.
    /// </summary>
    [Required(ErrorMessage = "Dispatch__SupportedSchemaVersions is required.")]
    public string SupportedSchemaVersions { get; set; } = string.Empty;

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
