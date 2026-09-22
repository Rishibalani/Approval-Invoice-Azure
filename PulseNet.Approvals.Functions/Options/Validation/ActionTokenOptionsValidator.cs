using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;

namespace PulseNet.Approvals.Functions.Options.Validation;

/// <summary>
/// The ActionToken checks DataAnnotations cannot express. Runs alongside the
/// attribute validation on ActionTokenOptions.
/// </summary>
internal sealed class ActionTokenOptionsValidator : IValidateOptions<ActionTokenOptions>
{
    private const string Section = ActionTokenOptions.SectionName;

    private readonly IConfiguration _configuration;

    public ActionTokenOptionsValidator(IConfiguration configuration) => _configuration = configuration;

    public ValidateOptionsResult Validate(string? name, ActionTokenOptions options)
    {
        var errors = new OptionsErrors(_configuration);

        // A bool cannot distinguish "false" from "missing" once bound. This
        // one gates Easy Auth enforcement, so it must be stated explicitly.
        errors.RequirePresent($"{Section}:RequireSignedInUser");

        return errors.ToResult();
    }
}
