using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;

namespace PulseNet.Approvals.Functions.Options.Validation;

/// <summary>
/// Collects validation failures for one options type, phrased in terms of the
/// app-setting key an operator actually has to set ("Channels__WhatsApp__AppSecret
/// is required"), not the C# property name.
///
/// WHY PRESENCE CHECKS AGAINST IConfiguration
///
/// A bool or enum property cannot tell "configured as false/Disabled" from
/// "missing" once bound - both are the type's zero value. For toggles that
/// matters: a missing RequireRejectionReason silently becoming false is
/// exactly the kind of fallback this app refuses. So those keys are checked
/// for presence in configuration itself.
/// </summary>
internal sealed class OptionsErrors
{
    private readonly List<string> _errors = [];
    private readonly IConfiguration _configuration;

    public OptionsErrors(IConfiguration configuration) => _configuration = configuration;

    /// <summary>"Channels:WhatsApp:AppSecret" -> "Channels__WhatsApp__AppSecret".</summary>
    public static string AppSetting(string configKey) => configKey.Replace(":", "__");

    /// <summary>The key must exist in configuration (used for bool and enum toggles).</summary>
    public void RequirePresent(string configKey)
    {
        if (_configuration[configKey] is null)
        {
            _errors.Add($"{AppSetting(configKey)} is required.");
        }
    }

    public void RequireValue(string? value, string configKey)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            _errors.Add($"{AppSetting(configKey)} is required.");
        }
    }

    public void RequireAbsoluteUrl(string? value, string configKey)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            _errors.Add($"{AppSetting(configKey)} is required.");
        }
        else if (!Uri.TryCreate(value, UriKind.Absolute, out var uri) ||
                 (uri.Scheme != Uri.UriSchemeHttps && uri.Scheme != Uri.UriSchemeHttp))
        {
            _errors.Add($"{AppSetting(configKey)} must be an absolute http(s) URL.");
        }
    }

    public void RequirePositive(int value, string configKey)
    {
        if (value <= 0)
        {
            _errors.Add($"{AppSetting(configKey)} must be > 0.");
        }
    }

    public void RequireRange(int value, int min, int max, string configKey)
    {
        if (value < min || value > max)
        {
            _errors.Add($"{AppSetting(configKey)} must be between {min} and {max}.");
        }
    }

    public void Add(string message) => _errors.Add(message);

    public ValidateOptionsResult ToResult() =>
        _errors.Count == 0 ? ValidateOptionsResult.Success : ValidateOptionsResult.Fail(_errors);
}
