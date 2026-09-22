using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Models;

namespace PulseNet.Approvals.Functions.Options.Validation;

/// <summary>
/// Validates the Channels section, including the nested Teams, Outlook and
/// WhatsApp classes that DataAnnotations would not recurse into.
///
/// ALWAYS REQUIRED
///   Each channel's DeliveryMode (and ActionMode for Teams/Outlook), the
///   fallback channel and currency, Outlook's RequireRejectionReason (read by
///   the Teams bot too), and the HttpClient timeouts - the typed clients are
///   built whenever their sender is resolved, enabled or not.
///
/// CONDITIONAL
///   Teams       toggles when enabled; WebhookUrl + PayloadMode in WorkflowWebhook
///   Outlook     transport and token settings when enabled
///   WhatsApp    every Cloud API setting when enabled
///   ActionEndpointBaseUrl when any enabled channel uses Link action mode
/// </summary>
internal sealed class ChannelOptionsValidator : IValidateOptions<ChannelOptions>
{
    private const string Section = ChannelOptions.SectionName;
    private const string Teams = Section + ":Teams";
    private const string Outlook = Section + ":Outlook";
    private const string WhatsApp = Section + ":WhatsApp";

    private readonly IConfiguration _configuration;

    public ChannelOptionsValidator(IConfiguration configuration) => _configuration = configuration;

    public ValidateOptionsResult Validate(string? name, ChannelOptions options)
    {
        var errors = new OptionsErrors(_configuration);

        errors.RequirePresent($"{Section}:FallbackChannel");
        errors.RequireValue(options.FallbackCurrencyCode, $"{Section}:FallbackCurrencyCode");
        errors.RequirePositive(options.MaxLinesOnCard, $"{Section}:MaxLinesOnCard");

        ValidateTeams(options.Teams, errors);
        ValidateOutlook(options.Outlook, errors);
        ValidateWhatsApp(options.WhatsApp, errors);

        var linkInUse =
            UsesLink(options.Teams.DeliveryMode, options.Teams.ActionMode) ||
            UsesLink(options.Outlook.DeliveryMode, options.Outlook.ActionMode) ||
            UsesLink(options.WhatsApp.DeliveryMode, options.WhatsApp.ActionMode);

        if (linkInUse || !string.IsNullOrWhiteSpace(options.ActionEndpointBaseUrl))
        {
            errors.RequireAbsoluteUrl(options.ActionEndpointBaseUrl, $"{Section}:ActionEndpointBaseUrl");
        }

        return errors.ToResult();
    }

    private static bool UsesLink(ChannelDeliveryMode delivery, ChannelActionMode action) =>
        delivery != ChannelDeliveryMode.Disabled && action == ChannelActionMode.Link;

    private static void ValidateTeams(TeamsChannelOptions teams, OptionsErrors errors)
    {
        errors.RequirePresent($"{Teams}:DeliveryMode");
        errors.RequirePresent($"{Teams}:ActionMode");
        errors.RequirePositive(teams.WebhookHttpTimeoutSeconds, $"{Teams}:WebhookHttpTimeoutSeconds");

        if (teams.DeliveryMode == ChannelDeliveryMode.Disabled)
        {
            return;
        }

        errors.RequirePresent($"{Teams}:UseFullWidthCard");
        errors.RequirePresent($"{Teams}:UseCardRefreshBlock");

        if (teams.DeliveryMode == ChannelDeliveryMode.WorkflowWebhook)
        {
            errors.RequireAbsoluteUrl(teams.WebhookUrl, $"{Teams}:WebhookUrl");
            errors.RequirePresent($"{Teams}:PayloadMode");
        }

        // Bot mode settings live in the TeamsBot section and are checked by
        // TeamsBotOptionsValidator.
    }

    private static void ValidateOutlook(OutlookChannelOptions outlook, OptionsErrors errors)
    {
        errors.RequirePresent($"{Outlook}:DeliveryMode");
        errors.RequirePresent($"{Outlook}:ActionMode");

        // Read by the Teams bot invoke as well - one rejection-reason policy,
        // not one per channel - so required whatever Outlook's own mode is.
        errors.RequirePresent($"{Outlook}:RequireRejectionReason");

        if (outlook.DeliveryMode == ChannelDeliveryMode.Disabled)
        {
            return;
        }

        errors.RequireValue(outlook.FromAddress, $"{Outlook}:FromAddress");
        errors.RequireValue(outlook.FromDisplayName, $"{Outlook}:FromDisplayName");
        errors.RequireValue(outlook.Transport, $"{Outlook}:Transport");
        errors.RequirePresent($"{Outlook}:ValidateInboundToken");
        errors.RequirePresent($"{Outlook}:RequireMailboxMatch");
        errors.RequireAbsoluteUrl(outlook.TokenMetadataUrl, $"{Outlook}:TokenMetadataUrl");
        errors.RequireValue(outlook.ValidTokenIssuers, $"{Outlook}:ValidTokenIssuers");

        if (outlook.ValidateInboundToken)
        {
            errors.RequireAbsoluteUrl(outlook.ExpectedTokenAudience, $"{Outlook}:ExpectedTokenAudience");
        }

        if (outlook.DeliveryMode == ChannelDeliveryMode.ActionableMessage)
        {
            errors.RequireValue(outlook.OriginatorId, $"{Outlook}:OriginatorId");
        }

        if (string.Equals(outlook.Transport, "Graph", StringComparison.OrdinalIgnoreCase))
        {
            errors.RequireValue(outlook.TenantId, $"{Outlook}:TenantId");
            errors.RequireValue(outlook.ClientId, $"{Outlook}:ClientId");
            errors.RequireValue(outlook.ClientSecret, $"{Outlook}:ClientSecret");
        }
        else if (string.Equals(outlook.Transport, "Acs", StringComparison.OrdinalIgnoreCase))
        {
            errors.RequireValue(outlook.AcsConnectionString, $"{Outlook}:AcsConnectionString");
        }
    }

    private static void ValidateWhatsApp(WhatsAppChannelOptions whatsApp, OptionsErrors errors)
    {
        errors.RequirePresent($"{WhatsApp}:DeliveryMode");
        errors.RequirePositive(whatsApp.HttpTimeoutSeconds, $"{WhatsApp}:HttpTimeoutSeconds");

        if (whatsApp.DeliveryMode == ChannelDeliveryMode.Disabled)
        {
            return;
        }

        errors.RequirePresent($"{WhatsApp}:ActionMode");
        errors.RequireValue(whatsApp.PhoneNumberId, $"{WhatsApp}:PhoneNumberId");
        errors.RequireValue(whatsApp.AccessToken, $"{WhatsApp}:AccessToken");
        errors.RequireValue(whatsApp.TemplateName, $"{WhatsApp}:TemplateName");
        errors.RequireValue(whatsApp.TemplateLanguage, $"{WhatsApp}:TemplateLanguage");
        errors.RequireValue(whatsApp.AppSecret, $"{WhatsApp}:AppSecret");
        errors.RequireValue(whatsApp.WebhookVerifyToken, $"{WhatsApp}:WebhookVerifyToken");
        errors.RequireAbsoluteUrl(whatsApp.GraphApiBaseUrl, $"{WhatsApp}:GraphApiBaseUrl");
        errors.RequireValue(whatsApp.ApiVersion, $"{WhatsApp}:ApiVersion");
        errors.RequireValue(whatsApp.PendingRejectionTable, $"{WhatsApp}:PendingRejectionTable");
        errors.RequireValue(whatsApp.InboundDedupeTable, $"{WhatsApp}:InboundDedupeTable");
        errors.RequirePositive(whatsApp.RejectionReasonTimeoutMinutes, $"{WhatsApp}:RejectionReasonTimeoutMinutes");
    }
}
