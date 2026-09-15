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
using PulseNet.Approvals.Functions.Cards;
using PulseNet.Approvals.Functions.Channels;
using PulseNet.Approvals.Functions.Options;
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
//  Set AppConfig__Endpoint to switch that on.
// ════════════════════════════════════════════════════════════════════════

var appConfigEndpoint = builder.Configuration["AppConfig:Endpoint"];

if (!string.IsNullOrWhiteSpace(appConfigEndpoint))
{
    builder.Configuration.AddAzureAppConfiguration(options =>
    {
        options.Connect(new Uri(appConfigEndpoint), new DefaultAzureCredential())
               .ConfigureKeyVault(kv => kv.SetCredential(new DefaultAzureCredential()))
               .ConfigureRefresh(refresh =>
                   refresh.Register("Dispatch:SigningSecret", refreshAll: false)
                          .SetRefreshInterval(TimeSpan.FromMinutes(1)));
    });
}

// ════════════════════════════════════════════════════════════════════════
//  OPTIONS
// ════════════════════════════════════════════════════════════════════════

builder.Services
    .AddOptions<DispatchOptions>()
    .Bind(builder.Configuration.GetSection(DispatchOptions.SectionName))
    .ValidateOnStart();

builder.Services
    .AddOptions<ChannelOptions>()
    .Bind(builder.Configuration.GetSection(ChannelOptions.SectionName))
    .ValidateOnStart();

builder.Services
    .AddOptions<ActionTokenOptions>()
    .Bind(builder.Configuration.GetSection(ActionTokenOptions.SectionName))
    .ValidateOnStart();

builder.Services
    .AddOptions<BusinessCentralOptions>()
    .Bind(builder.Configuration.GetSection(BusinessCentralOptions.SectionName))
    .ValidateOnStart();

builder.Services
    .AddOptions<TeamsBotOptions>()
    .Bind(builder.Configuration.GetSection(TeamsBotOptions.SectionName))
    .ValidateOnStart();

// ════════════════════════════════════════════════════════════════════════
//  STORAGE
//
//  Managed identity in Azure: set AzureWebJobsStorage__accountName and grant
//  the Function's identity Storage Table and Queue Data Contributor. The
//  emulator connection string during development.
// ════════════════════════════════════════════════════════════════════════

builder.Services.AddAzureClients(clients =>
{
    var storageConnection = builder.Configuration["AzureWebJobsStorage"];
    var storageAccountName = builder.Configuration["AzureWebJobsStorage:accountName"];

    if (!string.IsNullOrWhiteSpace(storageAccountName))
    {
        clients.AddTableServiceClient(
            new Uri($"https://{storageAccountName}.table.core.windows.net"));

        clients.AddQueueServiceClient(
            new Uri($"https://{storageAccountName}.queue.core.windows.net"))
            .ConfigureOptions(o => o.MessageEncoding = QueueMessageEncoding.Base64);

        clients.UseCredential(new DefaultAzureCredential());
    }
    else
    {
        clients.AddTableServiceClient(storageConnection);

        clients.AddQueueServiceClient(storageConnection)
            .ConfigureOptions(o => o.MessageEncoding = QueueMessageEncoding.Base64);
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
builder.Services.AddSingleton<OutlookCardBuilder>();
builder.Services.AddSingleton<ApprovalEmailBuilder>();

// ════════════════════════════════════════════════════════════════════════
//  CHANNEL: TEAMS VIA WEBHOOK
//
//  Registered against the concrete type first so the typed HttpClient is
//  honoured, then exposed through IChannelSender. Registering the interface
//  directly would build a second instance without the configured client.
// ════════════════════════════════════════════════════════════════════════

builder.Services.AddHttpClient<WorkflowWebhookSender>(client =>
{
    client.Timeout = TimeSpan.FromSeconds(30);
});

builder.Services.AddSingleton<IChannelSender>(sp =>
    sp.GetRequiredService<WorkflowWebhookSender>());

// ════════════════════════════════════════════════════════════════════════
//  CHANNEL: TEAMS VIA BOT
// ════════════════════════════════════════════════════════════════════════

builder.Services.AddHttpClient<BotConnectorClient>(client =>
{
    client.Timeout = TimeSpan.FromSeconds(30);
});

// Bulk directory reads and Teams app installation, used by the provisioning
// endpoint. Separate from BotConnectorClient because it is admin tooling
// rather than part of the dispatch path.
builder.Services.AddHttpClient<GraphDirectoryClient>(client =>
{
    // Generous: a directory page of 999 users on a slow tenant.
    client.Timeout = TimeSpan.FromSeconds(60);
});

builder.Services.AddSingleton<ConversationReferenceStore>();
builder.Services.AddSingleton<BotFrameworkTokenValidator>();
builder.Services.AddSingleton<TeamsBotSender>();

builder.Services.AddSingleton<IChannelSender>(sp =>
    sp.GetRequiredService<TeamsBotSender>());

// ════════════════════════════════════════════════════════════════════════
//  CHANNEL: OUTLOOK
//
//  Transport is chosen by configuration rather than compiled in, because which
//  one to use is a decision with cost and permission consequences that can
//  differ per environment.
//
//  Null is the default and returns FAILURE rather than pretending to have
//  sent. A transport that quietly lies would let the dispatcher mark the
//  channel delivered and skip the fallback, and nobody would learn the
//  approver was never told.
// ════════════════════════════════════════════════════════════════════════

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

builder.Services.AddSingleton<ActionableMessageTokenValidator>();

builder.Services.AddSingleton<IChannelSender, PlainEmailSender>();
builder.Services.AddSingleton<IChannelSender, OutlookActionableMessageSender>();

// ════════════════════════════════════════════════════════════════════════
//  CHANNEL: WHATSAPP
//
//  The only channel that needs conversational state. A quick-reply button
//  carries a payload and nothing else, so a rejection reason cannot arrive
//  with the tap - PendingRejectionStore holds the rejection between the tap
//  and the reply.
// ════════════════════════════════════════════════════════════════════════

builder.Services.AddHttpClient<WhatsAppClient>(client =>
{
    client.Timeout = TimeSpan.FromSeconds(30);
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

builder.Services.AddHttpClient<BusinessCentralClient>(client =>
{
    // Generous: an approval callback that times out leaves the approver
    // staring at a spinner with no idea whether it worked.
    client.Timeout = TimeSpan.FromSeconds(60);
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
