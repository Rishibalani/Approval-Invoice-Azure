using Azure.Data.Tables;
using Azure.Identity;
using Azure.Storage.Queues;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Builder;
using Microsoft.Extensions.Azure;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Cards;
using PulseNet.Approvals.Functions.Channels;
using PulseNet.Approvals.Functions.Options;
using PulseNet.Approvals.Functions.Options.Validation;
using PulseNet.Approvals.Functions.Security;
using PulseNet.Approvals.Functions.Services;

var builder = FunctionsApplication.CreateBuilder(args);

// ASP.NET Core integration. Needed so functions receive HttpRequest and can
// read the raw body as a string - the HMAC check depends on getting the exact
// bytes Business Central signed, which HttpRequestData makes awkward.
builder.ConfigureFunctionsWebApplication();

// ════════════════════════════════════════════════════════════════════════
//  CONFIGURATION
//
//  Local: local.settings.json.
//  Azure: App Settings, optionally fronted by Azure App Configuration with
//  Key Vault references so a rotated secret takes effect without a redeploy.
//  Set AppConfig__Endpoint to switch that on; AppConfig__SentinelKey and
//  AppConfig__RefreshIntervalSeconds are then required.
//
//  This runs before the options pipeline exists, so these three are read
//  straight from configuration rather than through IOptions.
// ════════════════════════════════════════════════════════════════════════

var appConfigEndpoint = builder.Configuration["AppConfig:Endpoint"];

if (!string.IsNullOrWhiteSpace(appConfigEndpoint))
{
    var sentinelKey = builder.Configuration["AppConfig:SentinelKey"];

    if (string.IsNullOrWhiteSpace(sentinelKey))
    {
        throw new InvalidOperationException(
            "AppConfig__SentinelKey is required when AppConfig__Endpoint is set.");
    }

    if (!int.TryParse(builder.Configuration["AppConfig:RefreshIntervalSeconds"], out var refreshSeconds) ||
        refreshSeconds <= 0)
    {
        throw new InvalidOperationException(
            "AppConfig__RefreshIntervalSeconds must be > 0 when AppConfig__Endpoint is set.");
    }

    builder.Configuration.AddAzureAppConfiguration(options =>
    {
        options.Connect(new Uri(appConfigEndpoint), new DefaultAzureCredential())
               .ConfigureKeyVault(kv => kv.SetCredential(new DefaultAzureCredential()))
               .ConfigureRefresh(refresh =>
                   refresh.Register(sentinelKey, refreshAll: false)
                          .SetRefreshInterval(TimeSpan.FromSeconds(refreshSeconds)));
    });
}

// ════════════════════════════════════════════════════════════════════════
//  OPTIONS
//
//  Strictly configuration-driven: no option class carries a code default for
//  a configurable value. DataAnnotations cover the always-required settings;
//  the IValidateOptions classes in Options/Validation cover the conditional
//  ones (WhatsApp only when enabled, TeamsBot only in Bot mode, ...) and the
//  bool/enum toggles whose absence a bound value cannot reveal. ValidateOnStart
//  means a missing setting stops the host at startup, naming the app setting,
//  instead of surfacing mid-approval.
// ════════════════════════════════════════════════════════════════════════

builder.Services
    .AddOptions<DispatchOptions>()
    .Bind(builder.Configuration.GetSection(DispatchOptions.SectionName))
    .ValidateDataAnnotations()
    .ValidateOnStart();

builder.Services
    .AddOptions<ChannelOptions>()
    .Bind(builder.Configuration.GetSection(ChannelOptions.SectionName))
    .ValidateDataAnnotations()
    .ValidateOnStart();

builder.Services.AddSingleton<IValidateOptions<ChannelOptions>, ChannelOptionsValidator>();

builder.Services
    .AddOptions<ActionTokenOptions>()
    .Bind(builder.Configuration.GetSection(ActionTokenOptions.SectionName))
    .ValidateDataAnnotations()
    .ValidateOnStart();

builder.Services.AddSingleton<IValidateOptions<ActionTokenOptions>, ActionTokenOptionsValidator>();

builder.Services
    .AddOptions<BusinessCentralOptions>()
    .Bind(builder.Configuration.GetSection(BusinessCentralOptions.SectionName))
    .ValidateDataAnnotations()
    .ValidateOnStart();

builder.Services
    .AddOptions<TeamsBotOptions>()
    .Bind(builder.Configuration.GetSection(TeamsBotOptions.SectionName))
    .ValidateDataAnnotations()
    .ValidateOnStart();

builder.Services.AddSingleton<IValidateOptions<TeamsBotOptions>, TeamsBotOptionsValidator>();

builder.Services
    .AddOptions<EntraOptions>()
    .Bind(builder.Configuration.GetSection(EntraOptions.SectionName))
    .ValidateDataAnnotations()
    .ValidateOnStart();

// ════════════════════════════════════════════════════════════════════════
//  STORAGE
//
//  Connection string (the emulator during development): set AzureWebJobsStorage.
//
//  Managed identity in Azure: leave AzureWebJobsStorage unset and set
//  AzureWebJobsStorage__tableServiceUri and AzureWebJobsStorage__queueServiceUri
//  - the host's own identity-based connection settings, so the queue trigger
//  and these clients read the same values - then grant the Function's
//  identity Storage Table and Queue Data Contributor.
// ════════════════════════════════════════════════════════════════════════

builder.Services.AddAzureClients(clients =>
{
    var storageConnection = builder.Configuration["AzureWebJobsStorage"];

    if (!string.IsNullOrWhiteSpace(storageConnection))
    {
        clients.AddTableServiceClient(storageConnection);

        clients.AddQueueServiceClient(storageConnection)
            .ConfigureOptions(o => o.MessageEncoding = QueueMessageEncoding.Base64);
    }
    else
    {
        var tableServiceUri = builder.Configuration["AzureWebJobsStorage:tableServiceUri"];
        var queueServiceUri = builder.Configuration["AzureWebJobsStorage:queueServiceUri"];

        if (string.IsNullOrWhiteSpace(tableServiceUri) || string.IsNullOrWhiteSpace(queueServiceUri))
        {
            throw new InvalidOperationException(
                "Storage is not configured. Set AzureWebJobsStorage, or both " +
                "AzureWebJobsStorage__tableServiceUri and AzureWebJobsStorage__queueServiceUri.");
        }

        clients.AddTableServiceClient(new Uri(tableServiceUri));

        clients.AddQueueServiceClient(new Uri(queueServiceUri))
            .ConfigureOptions(o => o.MessageEncoding = QueueMessageEncoding.Base64);

        clients.UseCredential(new DefaultAzureCredential());
    }
});

// ════════════════════════════════════════════════════════════════════════
//  SECURITY AND CORE SERVICES
// ════════════════════════════════════════════════════════════════════════

builder.Services.AddSingleton<RequestSignatureValidator>();
builder.Services.AddSingleton<IdempotencyStore>();
builder.Services.AddSingleton<ActionTokenService>();

// Token validation, identity check, Business Central call and nonce burn.
// Shared by the link endpoint, the Outlook endpoint and the bot invoke.
// Duplicating it is how the three drift apart, and the one that drifts is the
// one that stops burning nonces.
builder.Services.AddSingleton<ApprovalDecisionService>();

// ════════════════════════════════════════════════════════════════════════
//  CARD AND EMAIL CONSTRUCTION
//
//  All three read ApprovalCardViewModel, so the Teams card, the Outlook card
//  and the HTML fallback cannot disagree about what an invoice says.
// ════════════════════════════════════════════════════════════════════════

builder.Services.AddSingleton<ApprovalCardBuilder>();

// OutlookCardBuilder and ApprovalEmailBuilder are preserved but not
// registered - Business Central composes its own email now.
// builder.Services.AddSingleton<OutlookCardBuilder>();
// builder.Services.AddSingleton<ApprovalEmailBuilder>();

// ════════════════════════════════════════════════════════════════════════
//  CHANNEL: TEAMS VIA WEBHOOK
//
//  Registered against the concrete type first so the typed HttpClient is
//  honoured, then exposed through IChannelSender. Registering the interface
//  directly would build a second instance without the configured client.
//
//  Every HttpClient timeout below comes from its options class; the values
//  are validated > 0 at startup.
// ════════════════════════════════════════════════════════════════════════

builder.Services.AddHttpClient<WorkflowWebhookSender>((sp, client) =>
{
    client.Timeout = TimeSpan.FromSeconds(
        sp.GetRequiredService<IOptions<ChannelOptions>>().Value.Teams.WebhookHttpTimeoutSeconds);
});

builder.Services.AddSingleton<IChannelSender>(sp =>
    sp.GetRequiredService<WorkflowWebhookSender>());

// ════════════════════════════════════════════════════════════════════════
//  CHANNEL: TEAMS VIA BOT
// ════════════════════════════════════════════════════════════════════════

builder.Services.AddHttpClient<BotConnectorClient>((sp, client) =>
{
    client.Timeout = TimeSpan.FromSeconds(
        sp.GetRequiredService<IOptions<TeamsBotOptions>>().Value.HttpTimeoutSeconds);
});

// Bulk directory reads and Teams app installation, used by the provisioning
// endpoint. Separate from BotConnectorClient because it is admin tooling
// rather than part of the dispatch path.
builder.Services.AddHttpClient<GraphDirectoryClient>((sp, client) =>
{
    // Usually generous: a full directory page on a slow tenant.
    client.Timeout = TimeSpan.FromSeconds(
        sp.GetRequiredService<IOptions<TeamsBotOptions>>().Value.GraphHttpTimeoutSeconds);
});

builder.Services.AddSingleton<ConversationReferenceStore>();
builder.Services.AddSingleton<SentCardStore>();
builder.Services.AddSingleton<CardRefreshService>();
builder.Services.AddSingleton<BotFrameworkTokenValidator>();
builder.Services.AddSingleton<TeamsBotSender>();

builder.Services.AddSingleton<IChannelSender>(sp =>
    sp.GetRequiredService<TeamsBotSender>());

// ════════════════════════════════════════════════════════════════════════
//  CHANNEL: OUTLOOK - PRESERVED, NOT IN USE
//
//  Business Central now composes and sends approval emails itself, using its
//  own email module. That removed an app registration, a Global Administrator
//  consent, an Exchange application access policy and a shared mailbox - none
//  of which Business Central needs, because it already sends email and has
//  done since the day it was configured.
//
//  Azure still owns the ACTION endpoint. Buttons in a Business-Central-sent
//  email point at /api/approvals/act, and ApprovalActionFunction still
//  validates the token, burns the nonce, enforces the rejection reason and
//  calls back. Only composition and delivery moved.
//
//  The files under Channels, Cards, Services and Security are commented out
//  rather than deleted. To restore: uncomment them and this block.
// ════════════════════════════════════════════════════════════════════════

/*
builder.Services.AddHttpClient<GraphEmailTransport>(client =>
{
    client.Timeout = TimeSpan.FromSeconds(30);
});

builder.Services.AddSingleton<IEmailTransport>(sp =>
{
    var configured = builder.Configuration["Channels:Outlook:Transport"] ?? "None";

    return configured.ToLowerInvariant() switch
    {
        "graph" => sp.GetRequiredService<GraphEmailTransport>(),
        _ => new NullEmailTransport(
            sp.GetRequiredService<ILoggerFactory>().CreateLogger<NullEmailTransport>())
    };
});

builder.Services.AddSingleton<OutlookCardBuilder>();
builder.Services.AddSingleton<ApprovalEmailBuilder>();
builder.Services.AddSingleton<ActionableMessageTokenValidator>();

builder.Services.AddSingleton<IChannelSender, PlainEmailSender>();
builder.Services.AddSingleton<IChannelSender, OutlookActionableMessageSender>();
*/

// ════════════════════════════════════════════════════════════════════════
//  CHANNEL: WHATSAPP
//
//  The only channel that needs conversational state. A quick-reply button
//  carries a payload and nothing else, so a rejection reason cannot arrive
//  with the tap - PendingRejectionStore holds the rejection between the tap
//  and the reply.
// ════════════════════════════════════════════════════════════════════════

builder.Services.AddHttpClient<WhatsAppClient>((sp, client) =>
{
    client.Timeout = TimeSpan.FromSeconds(
        sp.GetRequiredService<IOptions<ChannelOptions>>().Value.WhatsApp.HttpTimeoutSeconds);
});

builder.Services.AddSingleton<PendingRejectionStore>();
builder.Services.AddSingleton<WhatsAppSignatureValidator>();
builder.Services.AddSingleton<WhatsAppTemplateSender>();

builder.Services.AddSingleton<IChannelSender>(sp =>
    sp.GetRequiredService<WhatsAppTemplateSender>());

// ════════════════════════════════════════════════════════════════════════
//  DISPATCH AND CALLBACK
// ════════════════════════════════════════════════════════════════════════

builder.Services.AddSingleton<ChannelDispatcher>();

builder.Services.AddHttpClient<BusinessCentralClient>((sp, client) =>
{
    // Usually generous: an approval callback that times out leaves the
    // approver staring at a spinner with no idea whether it worked.
    client.Timeout = TimeSpan.FromSeconds(
        sp.GetRequiredService<IOptions<BusinessCentralOptions>>().Value.HttpTimeoutSeconds);
});

// ════════════════════════════════════════════════════════════════════════
//  TELEMETRY
//
//  The correlation ID from Business Central flows into this, so one identifier
//  traces a request across both systems.
// ════════════════════════════════════════════════════════════════════════

builder.Services
    .AddApplicationInsightsTelemetryWorkerService()
    .ConfigureFunctionsApplicationInsights();

builder.Build().Run();
