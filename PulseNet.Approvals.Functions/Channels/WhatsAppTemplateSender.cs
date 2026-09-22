using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PulseNet.Approvals.Functions.Cards;
using PulseNet.Approvals.Functions.Models;
using PulseNet.Approvals.Functions.Options;
using PulseNet.Approvals.Functions.Security;
using PulseNet.Approvals.Functions.Services;

namespace PulseNet.Approvals.Functions.Channels;

/// <summary>
/// Sends the approval request as a WhatsApp template message.
///
/// HOW THIS DIFFERS FROM THE OTHER CHANNELS
///
/// Teams and Outlook render a card we build. WhatsApp renders a template Meta
/// approved, and all we supply are the values that fill its placeholders. The
/// layout, the wording and the button labels were fixed at approval time.
///
/// That has three consequences worth holding on to:
///
///   The field ORDER is load-bearing. Placeholders are positional - {{1}},
///   {{2}} - with no names. Swapping two arguments silently produces a message
///   showing the vendor where the amount should be, and Meta cannot catch it.
///
///   There are no document lines. A template has one body; ten lines of detail
///   would not fit and could not be laid out. The deep link carries that
///   weight instead.
///
///   Changing the wording means re-submitting to Meta. Unlike the other
///   channels, this text is not ours to edit freely.
///
/// IDENTITY ASSURANCE IS THE LOWEST OF THE THREE
///
/// A tap proves possession of a phone, nothing more - no Entra identity, no
/// sign-in, no Microsoft assertion about who is holding it. That is why
/// WhatsApp should carry the lowest approval ceiling of any channel, and why
/// Business Central sending canApproveInChannel=false here is a verdict to
/// respect rather than work around.
/// </summary>
public sealed class WhatsAppTemplateSender : IChannelSender
{
    /// <summary>Meta's cap on a quick-reply button payload. A Cloud API limit, not a setting.</summary>
    private const int MaxQuickReplyPayloadLength = 256;

    /// <summary>Meta's cap on a dynamic URL button suffix. A Cloud API limit, not a setting.</summary>
    private const int MaxUrlButtonSuffixLength = 2000;

    private readonly WhatsAppClient _client;
    private readonly ActionTokenService _tokenService;
    private readonly ChannelOptions _options;
    private readonly ILogger<WhatsAppTemplateSender> _logger;

    public ChannelDeliveryMode Mode => ChannelDeliveryMode.WhatsAppTemplate;
    public ApprovalChannel Channel => ApprovalChannel.WhatsApp;

    public WhatsAppTemplateSender(
        WhatsAppClient client,
        ActionTokenService tokenService,
        IOptions<ChannelOptions> options,
        ILogger<WhatsAppTemplateSender> logger)
    {
        _client = client;
        _tokenService = tokenService;
        _options = options.Value;
        _logger = logger;
    }

    public async Task<ChannelSendResult> SendAsync(
        ApprovalDispatchPayload payload,
        ChannelActionMode actionMode,
        string? approveUrl,
        string? rejectUrl,
        CancellationToken cancellationToken)
    {
        var mobile = payload.Approver.MobileNumber;

        if (string.IsNullOrWhiteSpace(mobile))
        {
            _logger.LogInformation(
                "No mobile number for {Approver}; skipping WhatsApp.",
                payload.Approver.UserId);

            return ChannelSendResult.Fail(Channel, "no_mobile_number");
        }

        // Consent is a legal gate, not a preference. Transmitting an
        // approver's name, a vendor and an amount to a third-party messaging
        // platform is processing that needs a recorded basis under GDPR and
        // the India DPDP Act. Business Central records it; this refuses
        // without it.
        if (!payload.Approver.ConsentGiven)
        {
            _logger.LogWarning(
                "No recorded WhatsApp consent for {Approver}; skipping.",
                payload.Approver.UserId);

            return ChannelSendResult.Fail(Channel, "no_whatsapp_consent");
        }

        var vm = ApprovalCardViewModel.From(payload, _options.FallbackCurrencyCode, _options.MaxLinesOnCard);

        try
        {
            // ---- Body parameters, in TEMPLATE ORDER --------------------
            //
            // This order must match the approved template exactly. See the
            // template guide - if the placeholders are re-ordered there, they
            // must be re-ordered here in the same commit.
            var bodyParameters = new List<string>
            {
                vm.TypeCaption,                                          // {{1}} Purchase invoice approval
                vm.DocumentNo,                                           // {{2}} INV26-00037
                vm.PartyName,                                            // {{3}} Catering Vendor
                vm.HeadlineAmount,                                       // {{4}} $11.30 CAD
                FirstFact(vm, "Due date") ?? "not set",                  // {{5}} 2026-10-03
                // RequesterLabel is nullable - name, then email, then user ID,
                // and all three can be empty. A template parameter may not be
                // null, and Meta rejects the whole send if one is.
                payload.Approval.RequesterLabel ?? "not recorded",        // {{6}} Rishi Balani
                vm.ChainContext ?? "Single approval"                     // {{7}} Approval 2 of 2
            };

            // ---- Buttons -----------------------------------------------
            var quickReplyPayloads = new List<string>();
            var canAct = actionMode != ChannelActionMode.NotifyOnly;

            if (canAct)
            {
                // Two tokens, two nonces. Burning Approve must not silently
                // disable Reject.
                //
                // The payload is capped at 256 characters by Meta, which is
                // exactly why the action token is a compact signed string
                // rather than a JWT.
                quickReplyPayloads.Add(_tokenService.Mint(
                    payload.Approval.ApprovalEntryNo, payload.Approver.Upn ?? mobile, ApprovalAction.Approve));

                quickReplyPayloads.Add(_tokenService.Mint(
                    payload.Approval.ApprovalEntryNo, payload.Approver.Upn ?? mobile, ApprovalAction.Reject));

                foreach (var p in quickReplyPayloads)
                {
                    if (p.Length > MaxQuickReplyPayloadLength)
                    {
                        // Should not happen - the token is about ninety
                        // characters - but a silent truncation by Meta would
                        // produce a token that fails validation with no clue
                        // why.
                        _logger.LogError(
                            "Action token is {Length} characters, over the {Limit} limit. Refusing to send.",
                            p.Length, MaxQuickReplyPayloadLength);

                        return ChannelSendResult.Fail(Channel, "token_too_long");
                    }
                }
            }

            // The URL button's base is fixed in the template; only a suffix is
            // dynamic. So the template holds the Business Central host and this
            // supplies the document-specific tail.
            var urlSuffix = ExtractDeepLinkSuffix(payload.Document.DeepLink);

            var result = await _client.SendTemplateAsync(
                mobile,
                bodyParameters,
                quickReplyPayloads,
                urlSuffix,
                cancellationToken);

            if (!result.Succeeded)
            {
                return ChannelSendResult.Fail(
                    Channel, result.FailureReason ?? "send_failed", result.IsTransient);
            }

            _logger.LogInformation(
                "WhatsApp template sent to {Approver} for {DocumentNo}, message {MessageId}.",
                payload.Approver.UserId, payload.Document.DocumentNo, result.MessageId);

            return ChannelSendResult.Ok(Channel, result.MessageId);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "WhatsApp sender threw for {DocumentNo}.", payload.Document.DocumentNo);
            return ChannelSendResult.Fail(Channel, "whatsapp_send_exception", transient: true);
        }
    }

    private static string? FirstFact(ApprovalCardViewModel vm, string title) =>
        vm.Facts.FirstOrDefault(f => f.Title == title)?.Value;

    /// <summary>
    /// The part of the deep link after the host, for a dynamic URL button.
    ///
    /// Meta fixes the base URL at template approval and permits only a suffix
    /// at send time - you cannot point the button somewhere else later, which
    /// is the whole point of approving it in the first place.
    /// </summary>
    private static string? ExtractDeepLinkSuffix(string? deepLink)
    {
        if (string.IsNullOrWhiteSpace(deepLink))
        {
            return null;
        }

        if (!Uri.TryCreate(deepLink, UriKind.Absolute, out var uri))
        {
            return null;
        }

        var suffix = uri.PathAndQuery.TrimStart('/');

        // Meta caps the suffix. Truncating a URL would produce a link that
        // opens the wrong record, so send nothing rather than something wrong -
        // the message body still names the document.
        return suffix.Length > MaxUrlButtonSuffixLength ? null : suffix;
    }
}
