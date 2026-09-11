using Azure.Data.Tables;
using Azure.Identity;
using Azure.Storage.Queues;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Builder;
using Microsoft.Extensions.Azure;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using PulseNet.Approvals.Functions.Options;
using PulseNet.Approvals.Functions.Security;
using PulseNet.Approvals.Functions.Services;
using PulseNet.Approvals.Functions.Cards;
using PulseNet.Approvals.Functions.Channels;

var builder = FunctionsApplication.CreateBuilder(args);

// ASP.NET Core integration. Needed so functions receive HttpRequest and can
// read the raw body as a string - the HMAC check depends on getting the exact
// bytes Business Central signed, which HttpRequestData makes awkward.
builder.ConfigureFunctionsWebApplication();

// ------------------------------------------------------------------------
//  Configuration
// ------------------------------------------------------------------------
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

// ------------------------------------------------------------------------
//  Storage Clients (Tables & Queues)
// ------------------------------------------------------------------------
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

//// Register solution security validators and services
//builder.Services.AddSingleton<RequestSignatureValidator>();
//builder.Services.AddSingleton<IdempotencyStore>();

builder.Services.AddSingleton<RequestSignatureValidator>();
builder.Services.AddSingleton<IdempotencyStore>();
builder.Services.AddSingleton<ActionTokenService>();
builder.Services.AddSingleton<ApprovalCardBuilder>();

// ------------------------------------------------------------------------
//  Channel senders
//
//  Registered against the interface, resolved by delivery mode at dispatch
//  time. Adding WhatsApp, or swapping the Teams webhook for a real bot, means
//  a new class and one more line here - never a change to ChannelDispatcher.
//
//  Typed HttpClients so the factory manages pooling and socket reuse. A
//  `new HttpClient()` per send exhausts ports under load.
// ------------------------------------------------------------------------
builder.Services.AddHttpClient<WorkflowWebhookSender>(client =>
{
    client.Timeout = TimeSpan.FromSeconds(30);
});

// Resolved via the concrete type so the typed HttpClient is honoured, then
// exposed through the interface for the dispatcher to enumerate.
builder.Services.AddSingleton<IChannelSender>(sp =>
    sp.GetRequiredService<WorkflowWebhookSender>());

builder.Services.AddSingleton<IChannelSender, PlainEmailSender>();

// ------------------------------------------------------------------------
//  Teams bot - 1:1 proactive messaging
//
//  Registered against the concrete type first so the typed HttpClient is
//  honoured, then exposed through IChannelSender. Registering the interface
//  directly would build a second instance without the configured client.
// ------------------------------------------------------------------------
builder.Services.AddHttpClient<BotConnectorClient>(client =>
{
    client.Timeout = TimeSpan.FromSeconds(30);
});

builder.Services.AddSingleton<ConversationReferenceStore>();
builder.Services.AddSingleton<BotFrameworkTokenValidator>();
builder.Services.AddSingleton<TeamsBotSender>();

builder.Services.AddSingleton<IChannelSender>(sp =>
    sp.GetRequiredService<TeamsBotSender>());

// Later stages add:
//   builder.Services.AddSingleton<IChannelSender, TeamsBotSender>();
//   builder.Services.AddSingleton<IChannelSender, ActionableMessageSender>();
//   builder.Services.AddSingleton<IChannelSender, WhatsAppTemplateSender>();

builder.Services.AddSingleton<ChannelDispatcher>();

builder.Services.AddHttpClient<BusinessCentralClient>(client =>
{
    // Generous: an approval callback that times out leaves the approver
    // staring at a spinner with no idea whether it worked.
    client.Timeout = TimeSpan.FromSeconds(60);
});

builder.Services
    .AddApplicationInsightsTelemetryWorkerService()
    .ConfigureFunctionsApplicationInsights();

builder.Build().Run();