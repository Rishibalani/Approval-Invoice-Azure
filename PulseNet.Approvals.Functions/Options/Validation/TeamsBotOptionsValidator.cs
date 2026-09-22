using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Models;

namespace PulseNet.Approvals.Functions.Options.Validation;

/// <summary>
/// TeamsBot settings are required only when Channels:Teams:DeliveryMode is Bot.
///
/// The HttpClient timeouts are the exception: they carry DataAnnotations on
/// TeamsBotOptions and are validated in every mode, because the typed clients
/// are constructed whenever the Teams sender or card refresh is resolved.
///
/// The delivery mode is read from configuration directly rather than through
/// IOptions&lt;ChannelOptions&gt;, so a problem in the Channels section is
/// reported once, by its own validator, instead of surfacing here as well.
/// </summary>
internal sealed class TeamsBotOptionsValidator : IValidateOptions<TeamsBotOptions>
{
    private const string Section = TeamsBotOptions.SectionName;
    private const string TeamsDeliveryModeKey = ChannelOptions.SectionName + ":Teams:DeliveryMode";

    private readonly IConfiguration _configuration;

    public TeamsBotOptionsValidator(IConfiguration configuration) => _configuration = configuration;

    public ValidateOptionsResult Validate(string? name, TeamsBotOptions options)
    {
        var botInUse =
            Enum.TryParse<ChannelDeliveryMode>(_configuration[TeamsDeliveryModeKey], ignoreCase: true, out var mode) &&
            mode == ChannelDeliveryMode.Bot;

        if (!botInUse)
        {
            return ValidateOptionsResult.Success;
        }

        var errors = new OptionsErrors(_configuration);

        errors.RequireValue(options.AppId, $"{Section}:AppId");
        errors.RequireValue(options.AppPassword, $"{Section}:AppPassword");
        errors.RequireValue(options.TenantId, $"{Section}:TenantId");
        errors.RequireAbsoluteUrl(options.DefaultServiceUrl, $"{Section}:DefaultServiceUrl");
        errors.RequireValue(options.ConversationTable, $"{Section}:ConversationTable");
        errors.RequireValue(options.SentCardTable, $"{Section}:SentCardTable");
        errors.RequirePresent($"{Section}:AttemptDirectConversation");

        errors.RequireValue(options.BotFrameworkScope, $"{Section}:BotFrameworkScope");
        errors.RequireAbsoluteUrl(options.OpenIdMetadataUrl, $"{Section}:OpenIdMetadataUrl");
        errors.RequireValue(options.ValidTokenIssuers, $"{Section}:ValidTokenIssuers");
        errors.RequirePositive(options.TokenClockSkewSeconds, $"{Section}:TokenClockSkewSeconds");

        errors.RequireAbsoluteUrl(options.GraphBaseUrl, $"{Section}:GraphBaseUrl");
        errors.RequireValue(options.GraphScope, $"{Section}:GraphScope");

        // Graph caps $top at 999 on /users.
        errors.RequireRange(options.GraphDirectoryPageSize, 1, 999, $"{Section}:GraphDirectoryPageSize");
        errors.RequirePositive(options.ProvisioningThrottleDelaySeconds, $"{Section}:ProvisioningThrottleDelaySeconds");

        return errors.ToResult();
    }
}
