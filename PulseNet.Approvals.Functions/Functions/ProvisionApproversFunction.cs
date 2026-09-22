using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Options;
using PulseNet.Approvals.Functions.Services;

namespace PulseNet.Approvals.Functions.Functions;

/// <summary>
/// Provisions approvers in bulk: resolves Entra object IDs and installs the
/// Teams app, for everyone at once rather than one at a time as dispatches
/// happen.
///
/// WHY THIS EXISTS AS A SEPARATE ENDPOINT
///
/// Doing it lazily during dispatch works, but it has two problems. The first
/// approval for each person pays a Graph round trip, and an approver who has
/// never installed the app falls through to email with no warning - which
/// looks like the Teams channel being broken rather than a provisioning gap.
///
/// Running this once after deployment, and again whenever approvers change,
/// means dispatch never has to think about it.
///
/// SAFE TO RE-RUN
///
/// Already-installed users return AlreadyInstalled, already-known object IDs
/// are skipped. Running it twice changes nothing, which matters because the
/// obvious reaction to a partial failure is to run it again.
///
/// ADMIN-TRIGGERED, NOT SCHEDULED
///
/// Deliberately HTTP with a function key rather than a timer. Installing an
/// app into people's Teams without anyone pressing a button is the kind of
/// thing that should be a decision, taken at a known moment, by someone who
/// can answer questions about it afterwards.
/// </summary>
public sealed class ProvisionApproversFunction
{
    private readonly GraphDirectoryClient _graph;
    private readonly BusinessCentralClient _bcClient;
    private readonly TeamsBotOptions _teamsOptions;
    private readonly ILogger<ProvisionApproversFunction> _logger;

    public ProvisionApproversFunction(
        GraphDirectoryClient graph,
        BusinessCentralClient bcClient,
        IOptions<TeamsBotOptions> teamsOptions,
        ILogger<ProvisionApproversFunction> logger)
    {
        _graph = graph;
        _bcClient = bcClient;
        _teamsOptions = teamsOptions.Value;
        _logger = logger;
    }

    [Function(nameof(ProvisionApprovers))]
    public async Task<IActionResult> ProvisionApprovers(
        [HttpTrigger(AuthorizationLevel.Function, "post", Route = "provisioning/approvers")] // "admin/..." is reserved by the Functions host
        HttpRequest req,
        CancellationToken cancellationToken)
    {
        // Default to a dry run. Someone exploring the endpoint should not
        // discover it by having an app appear in four hundred people's Teams.
        var apply = req.Query["apply"].FirstOrDefault() == "true";
        var install = req.Query["install"].FirstOrDefault() != "false";

        _logger.LogInformation(
            "Provisioning approvers. apply={Apply} install={Install}", apply, install);

        // ---- 1. Who are the approvers ---------------------------------
        var approvers = await _bcClient.GetApproverIdentitiesAsync(cancellationToken);

        if (approvers.Count == 0)
        {
            return new OkObjectResult(new
            {
                status = "nothing_to_do",
                message = "Business Central returned no approver identities. Run Import From Approval User Setup first."
            });
        }

        // ---- 2. The directory, in one pass ----------------------------
        var directory = await _graph.GetDirectoryAsync(cancellationToken);

        if (directory.Count == 0)
        {
            return new ObjectResult(new
            {
                status = "graph_unavailable",
                message = "Could not read the directory. Check User.Read.All is granted and consented."
            })
            { StatusCode = StatusCodes.Status503ServiceUnavailable };
        }

        // ---- 3. Catalogue app, if installing --------------------------
        string? catalogueAppId = null;

        if (install)
        {
            if (string.IsNullOrWhiteSpace(_teamsOptions.ManifestId))
            {
                return new BadRequestObjectResult(new
                {
                    status = "manifest_id_not_configured",
                    message = "TeamsBot__ManifestId must be set to the id from your Teams manifest before the app can be installed."
                });
            }

            catalogueAppId = await _graph.GetCatalogueAppIdAsync(_teamsOptions.ManifestId, cancellationToken);

            if (catalogueAppId is null)
            {
                return new BadRequestObjectResult(new
                {
                    status = "app_not_in_catalogue",
                    message = "No catalogue app matches that manifest id. Confirm the package was uploaded to Teams Admin Center and set to Allowed."
                });
            }
        }

        // ---- 4. Work through them -------------------------------------
        var report = new ProvisioningReport();

        foreach (var approver in approvers)
        {
            if (string.IsNullOrWhiteSpace(approver.Upn))
            {
                report.NoUpn.Add(approver.UserId);
                continue;
            }

            if (!directory.TryGetValue(approver.Upn, out var objectId))
            {
                // In Business Central but not in the directory. Usually a
                // disabled account or a UPN that has since changed.
                report.NotInDirectory.Add(approver.Upn);
                continue;
            }

            // ---- Object ID write-back ---------------------------------
            if (!string.Equals(approver.EntraObjectId, objectId, StringComparison.OrdinalIgnoreCase))
            {
                if (apply)
                {
                    var written = await _bcClient.SetEntraObjectIdAsync(
                        approver.UserId, objectId, cancellationToken);

                    if (written)
                    {
                        report.ObjectIdWritten.Add(approver.Upn);
                    }
                    else
                    {
                        report.WriteFailed.Add(approver.Upn);
                    }
                }
                else
                {
                    report.ObjectIdWouldWrite.Add(approver.Upn);
                }
            }
            else
            {
                report.ObjectIdAlreadySet.Add(approver.Upn);
            }

            // ---- Teams install ----------------------------------------
            if (!install || catalogueAppId is null)
            {
                continue;
            }

            if (!apply)
            {
                report.WouldInstall.Add(approver.Upn);
                continue;
            }

            var result = await _graph.InstallForUserAsync(objectId, catalogueAppId, cancellationToken);

            switch (result)
            {
                case InstallResult.Installed:
                    report.Installed.Add(approver.Upn);
                    break;
                case InstallResult.AlreadyInstalled:
                    report.AlreadyInstalled.Add(approver.Upn);
                    break;
                case InstallResult.NotEligible:
                    report.NoTeamsLicence.Add(approver.Upn);
                    break;
                case InstallResult.Throttled:
                    report.Throttled.Add(approver.Upn);
                    // Graph is pushing back. Pausing beats hammering it and
                    // having the rest of the run fail too.
                    await Task.Delay(
                        TimeSpan.FromSeconds(_teamsOptions.ProvisioningThrottleDelaySeconds),
                        cancellationToken);
                    break;
                default:
                    report.InstallFailed.Add(approver.Upn);
                    break;
            }
        }

        _logger.LogInformation("Provisioning finished. {Summary}", report.Summary);

        return new OkObjectResult(new
        {
            status = apply ? "applied" : "dry_run",
            hint = apply ? null : "Nothing was changed. Add ?apply=true to act on this.",
            approversInBc = approvers.Count,
            directoryUsers = directory.Count,
            report
        });
    }
}

/// <summary>
/// Outcomes by name rather than by count. A count tells you eleven approvers
/// were skipped; a list tells you which eleven, which is what somebody
/// actually needs in order to fix it.
/// </summary>
public sealed class ProvisioningReport
{
    public List<string> ObjectIdWritten { get; } = [];
    public List<string> ObjectIdWouldWrite { get; } = [];
    public List<string> ObjectIdAlreadySet { get; } = [];
    public List<string> Installed { get; } = [];
    public List<string> AlreadyInstalled { get; } = [];
    public List<string> WouldInstall { get; } = [];

    /// <summary>No Authentication Email on the Business Central user.</summary>
    public List<string> NoUpn { get; } = [];

    /// <summary>In Business Central but not in Entra - disabled, or the UPN changed.</summary>
    public List<string> NotInDirectory { get; } = [];

    public List<string> NoTeamsLicence { get; } = [];
    public List<string> Throttled { get; } = [];
    public List<string> InstallFailed { get; } = [];
    public List<string> WriteFailed { get; } = [];

    public string Summary =>
        $"objectIds written={ObjectIdWritten.Count} already={ObjectIdAlreadySet.Count}, " +
        $"installed={Installed.Count} already={AlreadyInstalled.Count}, " +
        $"noUpn={NoUpn.Count} notInDirectory={NotInDirectory.Count} " +
        $"noLicence={NoTeamsLicence.Count} failed={InstallFailed.Count + WriteFailed.Count}";
}
