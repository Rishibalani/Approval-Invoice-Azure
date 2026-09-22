using System.Text.Json.Serialization;

namespace PulseNet.Approvals.Functions.Models;

/// <summary>
/// The wire contract, mirroring what PN Approval Payload Builder emits in AL.
///
/// SCHEMA 1.1 ADDITIONS are all nullable or defaulted. A Business Central build
/// still emitting 1.0 deserialises cleanly and simply renders a thinner card,
/// so the AL extension and the Function can be deployed independently.
/// </summary>
public sealed record ApprovalDispatchPayload
{
    [JsonPropertyName("schemaVersion")] public string SchemaVersion { get; init; } = string.Empty;
    [JsonPropertyName("eventType")] public string EventType { get; init; } = string.Empty;
    [JsonPropertyName("eventId")] public string EventId { get; init; } = string.Empty;
    [JsonPropertyName("correlationId")] public string CorrelationId { get; init; } = string.Empty;
    [JsonPropertyName("occurredUtc")] public string OccurredUtc { get; init; } = string.Empty;

    [JsonPropertyName("source")] public SourceInfo Source { get; init; } = new();
    [JsonPropertyName("approval")] public ApprovalInfo Approval { get; init; } = new();
    [JsonPropertyName("document")] public DocumentInfo Document { get; init; } = new();
    [JsonPropertyName("approver")] public ApproverInfo Approver { get; init; } = new();
    [JsonPropertyName("policy")] public PolicyInfo Policy { get; init; } = new();
}

public sealed record SourceInfo
{
    [JsonPropertyName("system")] public string System { get; init; } = string.Empty;
    [JsonPropertyName("tenantId")] public string TenantId { get; init; } = string.Empty;
    [JsonPropertyName("environment")] public string Environment { get; init; } = string.Empty;
    [JsonPropertyName("companyName")] public string CompanyName { get; init; } = string.Empty;
    [JsonPropertyName("companyId")] public string CompanyId { get; init; } = string.Empty;

    // --- schema 1.1 ---

    /// <summary>
    /// The LCY code from General Ledger Setup. Business Central leaves
    /// Currency Code blank to mean "local", which is how "10.00" ended up on a
    /// card with no currency at all. Resolving blank needs this value.
    /// </summary>
    [JsonPropertyName("localCurrencyCode")] public string? LocalCurrencyCode { get; init; }

    /// <summary>Absolute https URL, 32x32 or larger. Its host must appear in the Teams manifest validDomains.</summary>
    [JsonPropertyName("brandIconUrl")] public string? BrandIconUrl { get; init; }
}

public sealed record ApprovalInfo
{
    [JsonPropertyName("approvalEntryNo")] public int ApprovalEntryNo { get; init; }
    [JsonPropertyName("sequenceNo")] public int SequenceNo { get; init; }
    [JsonPropertyName("recordId")] public string RecordId { get; init; } = string.Empty;
    [JsonPropertyName("status")] public string Status { get; init; } = string.Empty;
    [JsonPropertyName("dueDate")] public string? DueDate { get; init; }

    /// <summary>
    /// When the document was submitted for approval - as distinct from when
    /// the document itself was created, which can be weeks earlier.
    ///
    /// Business Central has been sending this; nothing bound it, so it was
    /// silently dropped. The card showed the document's creation date under a
    /// label that read like a submission date.
    /// </summary>
    [JsonPropertyName("sentForApprovalOn")] public string? SentForApprovalOn { get; init; }
    [JsonPropertyName("totalStepsInChain")] public int TotalStepsInChain { get; init; }
    [JsonPropertyName("openStepsRemaining")] public int OpenStepsRemaining { get; init; }
    [JsonPropertyName("isFinalStep")] public bool IsFinalStep { get; init; }
    [JsonPropertyName("requestedBy")] public string? RequestedBy { get; init; }

    // --- schema 1.1 ---

    /// <summary>Record SystemId. Drives the bookmark form of the deep link, which survives a renumbered document.</summary>
    [JsonPropertyName("recordSystemId")] public string? RecordSystemId { get; init; }

    /// <summary>Full name resolved from Approval Entry."Sender ID". RequestedBy alone is a user code, not a person.</summary>
    [JsonPropertyName("requestedByName")] public string? RequestedByName { get; init; }
    [JsonPropertyName("requestedByEmail")] public string? RequestedByEmail { get; init; }

    /// <summary>Set only when this entry reached the approver through BC's Delegate action.</summary>
    [JsonPropertyName("delegatedFromName")] public string? DelegatedFromName { get; init; }

    /// <summary>Best available label for the requester: name, then email, then the BC user code.</summary>
    [JsonIgnore]
    public string? RequesterLabel =>
        !string.IsNullOrWhiteSpace(RequestedByName) ? RequestedByName
        : !string.IsNullOrWhiteSpace(RequestedByEmail) ? RequestedByEmail
        : !string.IsNullOrWhiteSpace(RequestedBy) ? RequestedBy
        : null;
}

public sealed record DocumentInfo
{
    [JsonPropertyName("tableId")] public int TableId { get; init; }
    [JsonPropertyName("documentType")] public string DocumentType { get; init; } = string.Empty;
    [JsonPropertyName("documentNo")] public string DocumentNo { get; init; } = string.Empty;
    [JsonPropertyName("direction")] public string Direction { get; init; } = string.Empty;
    [JsonPropertyName("counterpartyNo")] public string? CounterpartyNo { get; init; }
    [JsonPropertyName("counterpartyName")] public string? CounterpartyName { get; init; }
    [JsonPropertyName("externalDocumentNo")] public string? ExternalDocumentNo { get; init; }
    [JsonPropertyName("amount")] public decimal Amount { get; init; }
    [JsonPropertyName("amountLcy")] public decimal AmountLcy { get; init; }
    [JsonPropertyName("currencyCode")] public string? CurrencyCode { get; init; }
    [JsonPropertyName("documentDate")] public string? DocumentDate { get; init; }
    [JsonPropertyName("dueDate")] public string? DueDate { get; init; }
    [JsonPropertyName("postingDate")] public string? PostingDate { get; init; }
    [JsonPropertyName("deepLink")] public string? DeepLink { get; init; }
    [JsonPropertyName("documentFound")] public bool DocumentFound { get; init; }
    [JsonPropertyName("lines")] public IReadOnlyList<DocumentLine>? Lines { get; init; }

    // --- schema 1.1 ---

    /// <summary>Human caption: "Purchase Invoice". Direction alone cannot distinguish an invoice from an order.</summary>
    [JsonPropertyName("documentTypeCaption")] public string? DocumentTypeCaption { get; init; }

    /// <summary>Pay-to / Bill-to party. Differs from the buy-from party often enough to be worth showing.</summary>
    [JsonPropertyName("payToNo")] public string? PayToNo { get; init; }
    [JsonPropertyName("payToName")] public string? PayToName { get; init; }

    /// <summary><see cref="Amount"/> is BC's net figure. These two make the tax explicit.</summary>
    [JsonPropertyName("amountExclTax")] public decimal AmountExclTax { get; init; }
    [JsonPropertyName("amountInclTax")] public decimal AmountInclTax { get; init; }
    [JsonPropertyName("taxAmount")] public decimal TaxAmount { get; init; }
    [JsonPropertyName("invoiceDiscountAmount")] public decimal InvoiceDiscountAmount { get; init; }

    [JsonPropertyName("createdByName")] public string? CreatedByName { get; init; }
    [JsonPropertyName("createdUtc")] public string? CreatedUtc { get; init; }

    /// <summary>Pre-joined "DEPT — Sales". Sending the code alone makes the approver decode it.</summary>
    [JsonPropertyName("dimension1Display")] public string? Dimension1Display { get; init; }
    [JsonPropertyName("dimension2Display")] public string? Dimension2Display { get; init; }

    /// <summary>Count of ALL lines on the document. Lines is capped for card size; this is not.</summary>
    [JsonPropertyName("totalLineCount")] public int TotalLineCount { get; init; }
    [JsonPropertyName("attachmentCount")] public int AttachmentCount { get; init; }

    /// <summary>
    /// False on a 1.0 payload, where AL sends no tax breakdown at all and every
    /// figure defaults to zero. Distinguishes "not sent" from "genuinely zero".
    /// </summary>
    [JsonIgnore] public bool HasTaxBreakdown => AmountInclTax != 0m || AmountExclTax != 0m;

    /// <summary>
    /// True when the pay-to party is not the buy-from party. Showing both when
    /// they match just makes the card longer.
    /// </summary>
    [JsonIgnore]
    public bool PayToDiffers =>
        !string.IsNullOrWhiteSpace(PayToName) &&
        !string.Equals(PayToName, CounterpartyName, StringComparison.OrdinalIgnoreCase);

}

public sealed record DocumentLine
{
    [JsonPropertyName("lineNo")] public int LineNo { get; init; }
    [JsonPropertyName("description")] public string Description { get; init; } = string.Empty;
    [JsonPropertyName("quantity")] public decimal Quantity { get; init; }
    [JsonPropertyName("lineAmount")] public decimal LineAmount { get; init; }

    // --- schema 1.1 ---
    [JsonPropertyName("unitOfMeasure")] public string? UnitOfMeasure { get; init; }
}

public sealed record ApproverInfo
{
    [JsonPropertyName("userId")] public string UserId { get; init; } = string.Empty;
    [JsonPropertyName("userSecurityId")] public string UserSecurityId { get; init; } = string.Empty;
    [JsonPropertyName("displayName")] public string? DisplayName { get; init; }
    [JsonPropertyName("upn")] public string? Upn { get; init; }
    [JsonPropertyName("entraObjectId")] public string? EntraObjectId { get; init; }
    [JsonPropertyName("suspended")] public bool Suspended { get; init; }
    /// <summary>
    /// Channel Business Central wants tried when every requested one fails.
    /// Null when the payload omits it; ChannelDispatcher then uses the
    /// configured Channels:FallbackChannel rather than a value baked in here.
    /// </summary>
    [JsonPropertyName("fallbackChannel")] public string? FallbackChannel { get; init; }
    [JsonPropertyName("channels")] public IReadOnlyList<string> Channels { get; init; } = [];

    /// <summary>E.164 mobile number. WhatsApp only.</summary>
    [JsonPropertyName("mobileNumber")] public string? MobileNumber { get; init; }

    /// <summary>
    /// Whether this approver has a recorded consent for WhatsApp processing.
    ///
    /// Defaults FALSE, and deliberately so. Transmitting a name, a vendor and
    /// an amount to a third-party messaging platform needs a recorded basis
    /// under GDPR and the India DPDP Act - a missing flag must mean "no", not
    /// "probably fine".
    /// </summary>
    [JsonPropertyName("consentGiven")] public bool ConsentGiven { get; init; }

    // --- schema 1.1 ---

    /// <summary>
    /// From User Setup."Substitute". Null hides the Delegate button entirely,
    /// which is the correct behaviour: a button that always errors is worse
    /// than no button.
    /// </summary>
    [JsonPropertyName("substituteUserId")] public string? SubstituteUserId { get; init; }
    [JsonPropertyName("substituteName")] public string? SubstituteName { get; init; }
    [JsonPropertyName("substituteUpn")] public string? SubstituteUpn { get; init; }
}

public sealed record PolicyInfo
{
    /// <summary>
    /// Business Central has already made this decision. When false the card
    /// must be notify-only with a deep link; the Function does not re-derive it
    /// from the amount, because financial policy belongs in the system of record.
    /// </summary>
    [JsonPropertyName("canApproveInChannel")] public bool CanApproveInChannel { get; init; }

    [JsonPropertyName("suppressionReasons")] public IReadOnlyList<string> SuppressionReasons { get; init; } = [];
    [JsonPropertyName("highValue")] public bool HighValue { get; init; }
    [JsonPropertyName("bankDetailsChanged")] public bool BankDetailsChanged { get; init; }
    /// <summary>
    /// Business Central's view of the button lifetime, informational only. The
    /// TTL actually minted into tokens is ActionToken:TtlMinutes. 0 when link
    /// expiry is switched off.
    /// </summary>
    [JsonPropertyName("actionTokenTtlMinutes")] public int ActionTokenTtlMinutes { get; init; }

    /// <summary>
    /// Business Central's global "Action Link Expiry Enabled" switch. False:
    /// buttons are minted without an expiry and stay usable until the approval
    /// is decided. Null only on payloads from an extension older than 1.0.0.4,
    /// which always expired its links - see ActionTokensExpire.
    /// </summary>
    [JsonPropertyName("actionTokenExpiryEnabled")] public bool? ActionTokenExpiryEnabled { get; init; }

    /// <summary>
    /// Whether tokens minted for this payload expire. A payload that predates
    /// the switch keeps the behaviour it was built with (expiring) - the safe
    /// reading of a missing security field is the stricter one.
    /// </summary>
    [JsonIgnore] public bool ActionTokensExpire => ActionTokenExpiryEnabled ?? true;
    [JsonPropertyName("requiresSignedInApproval")] public bool RequiresSignedInApproval { get; init; }

    /// <summary>Set false to hide Delegate even when a substitute exists.</summary>
    [JsonPropertyName("canDelegateInChannel")] public bool CanDelegateInChannel { get; init; } = true;

    [JsonPropertyName("perChannel")]
    public Dictionary<string, ChannelPolicy>? PerChannel { get; init; }

    public ChannelPolicy ResolveFor(string channelName)
    {
        if (!CanApproveInChannel)
        {
            return new ChannelPolicy { CanApprove = false, Reasons = SuppressionReasons };
        }

        if (PerChannel is not null && PerChannel.TryGetValue(channelName, out var specific))
        {
            return specific;
        }

        return new ChannelPolicy { CanApprove = true, Reasons = [] };
    }
}

public sealed record ChannelPolicy
{
    [JsonPropertyName("canApprove")] public bool CanApprove { get; init; }
    [JsonPropertyName("reasons")] public IReadOnlyList<string> Reasons { get; init; } = [];
}
