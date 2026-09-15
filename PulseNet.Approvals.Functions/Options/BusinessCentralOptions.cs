namespace PulseNet.Approvals.Functions.Options;

/// <summary>
/// Credentials for calling BACK into Business Central - the Phase 2 direction.
/// These belong to Entra app registration B, the one with API.ReadWrite.All
/// and admin consent, whose BC application user holds Approval Administrator.
/// </summary>
public sealed class BusinessCentralOptions
{
    public const string SectionName = "BusinessCentral";

    public string TenantId { get; set; } = string.Empty;
    public string ClientId { get; set; } = string.Empty;
    public string ClientSecret { get; set; } = string.Empty;

    /// <summary>Environment name, e.g. SANDBOX or PRODUCTION.</summary>
    public string EnvironmentName { get; set; } = string.Empty;

    /// <summary>Company GUID from /api/v2.0/companies.</summary>
    public string CompanyId { get; set; } = string.Empty;

    public string ApiPublisher { get; set; } = "pulsenet";
    public string ApiGroup { get; set; } = "approvals";
    public string ApiVersion { get; set; } = "v1.0";

    public string BaseUrl => $"https://api.businesscentral.dynamics.com/v2.0/{TenantId}/{EnvironmentName}";

    /// <summary>
    /// The approver identity API. Used to read who the approvers are and to
    /// write back the Entra object ID resolved from Graph.
    /// </summary>
    public string IdentitiesUrl =>
        $"{BaseUrl}/api/{ApiPublisher}/{ApiGroup}/{ApiVersion}/companies({CompanyId})/pnApproverIdentities";

    public string ApprovalEntriesUrl =>
        $"{BaseUrl}/api/{ApiPublisher}/{ApiGroup}/{ApiVersion}/companies({CompanyId})/pnApprovalEntries";

    public const string Scope = "https://api.businesscentral.dynamics.com/.default";
}
