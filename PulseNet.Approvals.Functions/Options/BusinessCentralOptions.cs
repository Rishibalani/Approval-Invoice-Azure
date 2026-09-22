using System.ComponentModel.DataAnnotations;

namespace PulseNet.Approvals.Functions.Options;

/// <summary>
/// Credentials for calling BACK into Business Central - the Phase 2 direction.
/// These belong to Entra app registration B, the one with API.ReadWrite.All
/// and admin consent, whose BC application user holds Approval Administrator.
///
/// Always required: every setting here is validated at startup, whichever
/// channels are enabled.
/// </summary>
public sealed class BusinessCentralOptions
{
    public const string SectionName = "BusinessCentral";

    [Required(ErrorMessage = "BusinessCentral__TenantId is required.")]
    public string TenantId { get; set; } = string.Empty;

    [Required(ErrorMessage = "BusinessCentral__ClientId is required.")]
    public string ClientId { get; set; } = string.Empty;

    [Required(ErrorMessage = "BusinessCentral__ClientSecret is required.")]
    public string ClientSecret { get; set; } = string.Empty;

    /// <summary>Environment name, e.g. SANDBOX or PRODUCTION.</summary>
    [Required(ErrorMessage = "BusinessCentral__EnvironmentName is required.")]
    public string EnvironmentName { get; set; } = string.Empty;

    /// <summary>Company GUID from /api/v2.0/companies.</summary>
    [Required(ErrorMessage = "BusinessCentral__CompanyId is required.")]
    public string CompanyId { get; set; } = string.Empty;

    [Required(ErrorMessage = "BusinessCentral__ApiPublisher is required.")]
    public string ApiPublisher { get; set; } = string.Empty;

    [Required(ErrorMessage = "BusinessCentral__ApiGroup is required.")]
    public string ApiGroup { get; set; } = string.Empty;

    [Required(ErrorMessage = "BusinessCentral__ApiVersion is required.")]
    public string ApiVersion { get; set; } = string.Empty;

    /// <summary>
    /// Business Central API root, up to and including the API version segment
    /// but before the tenant, e.g. https://api.businesscentral.dynamics.com/v2.0
    /// </summary>
    [Required(ErrorMessage = "BusinessCentral__ApiBaseUrl is required.")]
    [Url(ErrorMessage = "BusinessCentral__ApiBaseUrl must be an absolute URL.")]
    public string ApiBaseUrl { get; set; } = string.Empty;

    /// <summary>
    /// OAuth scope requested for the client-credentials token, e.g.
    /// https://api.businesscentral.dynamics.com/.default
    /// </summary>
    [Required(ErrorMessage = "BusinessCentral__Scope is required.")]
    public string Scope { get; set; } = string.Empty;

    /// <summary>
    /// HttpClient timeout for calls into Business Central. Generous: an
    /// approval callback that times out leaves the approver staring at a
    /// spinner with no idea whether it worked.
    /// </summary>
    [Range(1, int.MaxValue, ErrorMessage = "BusinessCentral__HttpTimeoutSeconds must be > 0.")]
    public int HttpTimeoutSeconds { get; set; }

    public string BaseUrl => $"{ApiBaseUrl.TrimEnd('/')}/{TenantId}/{EnvironmentName}";

    /// <summary>
    /// The approver identity API. Used to read who the approvers are and to
    /// write back the Entra object ID resolved from Graph.
    /// </summary>
    public string IdentitiesUrl =>
        $"{BaseUrl}/api/{ApiPublisher}/{ApiGroup}/{ApiVersion}/companies({CompanyId})/pnApproverIdentities";

    public string ApprovalEntriesUrl =>
        $"{BaseUrl}/api/{ApiPublisher}/{ApiGroup}/{ApiVersion}/companies({CompanyId})/pnApprovalEntries";
}
