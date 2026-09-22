using System.Globalization;
using PulseNet.Approvals.Functions.Models;

namespace PulseNet.Approvals.Functions.Cards;

/// <summary>
/// Everything the card needs, already formatted as strings.
///
/// WHY THIS EXISTS
///
/// The old builder formatted inline and used "-" as the placeholder for a
/// missing value. Adaptive Cards renders FactSet values as markdown, and a
/// lone "-" is markdown for an empty bullet - which is exactly the floating
/// bullet under "Requested by" on the current card.
///
/// The rule here: a value that is missing produces NO ROW. There is no
/// placeholder character, because every placeholder is a markdown hazard.
///
/// The money fields on DocumentInfo are plain decimals, not nullables, so the
/// Outlook builders can use them without unwrapping. "Not sent" is therefore
/// indistinguishable from zero on the raw field, and DocumentInfo.HasTaxBreakdown
/// exists to make that distinction instead.
/// </summary>
public sealed record ApprovalCardViewModel
{
    public string? BrandIconUrl { get; init; }
    public string TypeCaption { get; init; } = "Approval request";
    public string DocumentNo { get; init; } = string.Empty;
    public string PartyName { get; init; } = string.Empty;
    public string HeadlineAmount { get; init; } = string.Empty;
    public string? CreatedLine { get; init; }
    public string? DelegatedFromName { get; init; }
    public string? ChainContext { get; init; }
    public string? AttachmentNote { get; init; }
    public string? SubstituteName { get; init; }

    public IReadOnlyList<CardFact> Facts { get; init; } = [];
    public IReadOnlyList<CardLine> Lines { get; init; } = [];
    public int HiddenLineCount { get; init; }

    public sealed record CardFact(string Title, string Value);
    public sealed record CardLine(string Description, string Quantity, string UnitOfMeasure, string Amount);

    private static readonly CultureInfo Ci = CultureInfo.InvariantCulture;

    /// <param name="p">The dispatch payload.</param>
    /// <param name="fallbackCurrencyCode">
    /// Channels:FallbackCurrencyCode - used only when Business Central sends
    /// neither a document Currency Code nor an LCY code.
    /// </param>
    /// <param name="maxLinesOnCard">
    /// Channels:MaxLinesOnCard - document lines rendered before "+N more".
    /// Teams rejects cards over 28 KB, so keep this modest.
    /// </param>
    public static ApprovalCardViewModel From(ApprovalDispatchPayload p, string fallbackCurrencyCode, int maxLinesOnCard)
    {
        var doc = p.Document;
        var appr = p.Approval;

        // Blank Currency Code in BC means LCY. Resolve it, or the card shows a
        // bare number and the approver has to guess the currency.
        var currency = string.IsNullOrWhiteSpace(doc.CurrencyCode)
            ? p.Source.LocalCurrencyCode
            : doc.CurrencyCode;

        // On a 1.0 payload there is no tax breakdown, so both figures collapse
        // back to the single Amount that Business Central has always sent.
        var inclTax = doc.HasTaxBreakdown ? doc.AmountInclTax : doc.Amount;
        var exclTax = doc.HasTaxBreakdown ? doc.AmountExclTax : doc.Amount;

        var shownLineCount = Math.Min(doc.Lines?.Count ?? 0, maxLinesOnCard);
        var trueLineCount = doc.TotalLineCount > 0 ? doc.TotalLineCount : doc.Lines?.Count ?? 0;

        return new ApprovalCardViewModel
        {
            BrandIconUrl = Trim(p.Source.BrandIconUrl),
            TypeCaption = Trim(doc.DocumentTypeCaption)
                          ?? (doc.Direction == "Payable" ? "Purchase invoice approval" : "Sales invoice approval"),
            DocumentNo = doc.DocumentNo,
            PartyName = Trim(doc.CounterpartyName) ?? Trim(doc.CounterpartyNo) ?? "Unknown party",
            HeadlineAmount = Money(inclTax, currency, fallbackCurrencyCode),

            CreatedLine = BuildCreatedLine(doc),
            DelegatedFromName = Trim(appr.DelegatedFromName),
            ChainContext = BuildChainContext(appr),
            AttachmentNote = doc.AttachmentCount > 0
                ? $"{doc.AttachmentCount} attachment(s) on this document."
                : null,
            SubstituteName = Trim(p.Approver.SubstituteName),

            Facts = BuildFacts(p, currency, exclTax, inclTax, fallbackCurrencyCode),
            Lines = BuildLines(doc, currency, fallbackCurrencyCode, maxLinesOnCard),
            HiddenLineCount = Math.Max(0, trueLineCount - shownLineCount)
        };
    }

    // ------------------------------------------------------------------

    private static IReadOnlyList<CardFact> BuildFacts(
        ApprovalDispatchPayload p, string? currency, decimal exclTax, decimal inclTax, string fallbackCurrencyCode)
    {
        var doc = p.Document;
        var facts = new List<CardFact>();

        void Add(string title, string? value)
        {
            var clean = Trim(value);
            if (clean is not null) facts.Add(new CardFact(title, clean));
        }

        Add("Document", doc.DocumentNo);
        Add("Vendor", doc.CounterpartyName);

        // Only when it actually differs - repeating the same name twice makes
        // the card longer without making it more informative.
        if (doc.PayToDiffers) Add("Pay-to", doc.PayToName);

        Add("Their reference", doc.ExternalDocumentNo);

        // Same rule as the Outlook email (PN Approval Email Sender): the tax
        // breakdown when there is tax, otherwise a single Amount row.
        if (doc.HasTaxBreakdown && exclTax != inclTax)
        {
            Add("Amount excl. tax", Money(exclTax, currency, fallbackCurrencyCode));
            if (doc.TaxAmount > 0) Add("Tax", Money(doc.TaxAmount, currency, fallbackCurrencyCode));
            Add("Amount incl. tax", Money(inclTax, currency, fallbackCurrencyCode));
        }
        else
        {
            Add("Amount", Money(inclTax, currency, fallbackCurrencyCode));
        }

        // Local value is only interesting on a foreign-currency document.
        if (!string.IsNullOrWhiteSpace(doc.CurrencyCode) && doc.Amount != doc.AmountLcy)
        {
            Add("Local value", Money(doc.AmountLcy, p.Source.LocalCurrencyCode, fallbackCurrencyCode));
        }

        Add("Document date", Date(doc.DocumentDate));
        Add("Posting date", Date(doc.PostingDate));
        Add("Due date", Date(doc.DueDate));
        Add("Dimension", doc.Dimension1Display);
        Add("Cost centre", doc.Dimension2Display);

        // Name, then email, then the BC user code. When all three are empty the
        // row is dropped entirely - this is the floating-bullet fix.
        Add("Requested by", p.Approval.RequesterLabel);

        // WHEN it was submitted, not when the document was created. Those can
        // be weeks apart, and an approver deciding whether something is urgent
        // needs the submission date - an invoice raised in March and submitted
        // yesterday is not a month old in any sense that matters here.
        Add("Submitted", Timestamp(p.Approval.SentForApprovalOn));

        // "Respond by" removed. It duplicated the invoice Due date closely
        // enough to be read as the same thing, and an approver comparing two
        // dates that mean different things is worse served than one shown a
        // single date that matters.

        return facts;
    }

    private static IReadOnlyList<CardLine> BuildLines(
        DocumentInfo doc, string? currency, string fallbackCurrencyCode, int maxLinesOnCard)
    {
        if (doc.Lines is null || doc.Lines.Count == 0) return [];

        // Quantity and unit of measure in separate columns; amount carries the
        // currency, as in the Outlook email.
        return doc.Lines
            .Take(maxLinesOnCard)
            .Select(l => new CardLine(
                Description: string.IsNullOrWhiteSpace(l.Description) ? "(no description)" : l.Description,
                Quantity: l.Quantity.ToString("0.#####", Ci),
                UnitOfMeasure: Trim(l.UnitOfMeasure) ?? string.Empty,
                Amount: Money(l.LineAmount, currency, fallbackCurrencyCode)))
            .ToList();
    }

    /// <summary>
    /// A UTC instant rendered for reading. Returns null on anything
    /// unparseable, so the row is dropped rather than showing a raw ISO string
    /// to somebody trying to approve an invoice.
    /// </summary>
    private static string? Timestamp(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;

        return DateTimeOffset.TryParse(value, Ci, DateTimeStyles.AssumeUniversal, out var dt)
            ? dt.UtcDateTime.ToString("MM/dd/yy HH:mm", Ci) + " UTC"
            : null;
    }

    private static string? BuildCreatedLine(DocumentInfo doc)
    {
        var name = Trim(doc.CreatedByName);
        var when = DateTimeOffset.TryParse(doc.CreatedUtc, Ci, DateTimeStyles.AssumeUniversal, out var dt)
            ? dt.UtcDateTime.ToString("MM/dd/yy HH:mm", Ci) + " UTC"
            : null;

        return (name, when) switch
        {
            (null, null) => null,
            (not null, null) => $"Created by **{name}**",
            (null, not null) => $"Created {when}",
            _ => $"Created by **{name}** · {when}"
        };
    }

    private static string? BuildChainContext(ApprovalInfo a)
    {
        if (a.TotalStepsInChain <= 1) return null;

        return a.IsFinalStep
            ? $"Approval {a.SequenceNo} of {a.TotalStepsInChain}. This is the final approval - approving releases the document."
            : $"Approval {a.SequenceNo} of {a.TotalStepsInChain}. Further approval is required after yours.";
    }

    // ------------------------------------------------------------------

    private static string? Date(string? isoDate)
    {
        if (string.IsNullOrWhiteSpace(isoDate)) return null;
        return DateTime.TryParse(isoDate, Ci, DateTimeStyles.None, out var d)
            ? d.ToString("MM/dd/yy", Ci)
            : isoDate;
    }

    private static string? Trim(string? s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();

    /// <summary>
    /// Symbols for the currencies actually in use. Deliberately not a
    /// CultureInfo lookup — culture maps country to currency, which breaks the
    /// moment one company transacts in two currencies.
    /// </summary>
    private static readonly Dictionary<string, string> CurrencySymbols = new(StringComparer.OrdinalIgnoreCase)
    {
        ["CAD"] = "$",
        ["USD"] = "$",
        ["EUR"] = "€",
        ["GBP"] = "£",
        ["INR"] = "₹",
        ["AUD"] = "$",
        ["JPY"] = "¥"
    };

    /// <summary>
    /// fallbackCurrencyCode (Channels:FallbackCurrencyCode) is the last resort
    /// when Business Central sends neither a document Currency Code nor an LCY
    /// code. Wrong is better than absent here: a bare "10.00" tells the
    /// approver nothing, and the configured value is the company's home
    /// currency in every realistic case.
    /// </summary>
    private static string Money(decimal amount, string? currencyCode, string fallbackCurrencyCode)
    {
        var code = string.IsNullOrWhiteSpace(currencyCode) ? fallbackCurrencyCode.Trim() : currencyCode.Trim();
        var n = amount.ToString("N2", Ci);

        // Symbol AND code. "$10.00" is ambiguous across CAD, USD and AUD, and on
        // a purchase invoice that ambiguity is expensive.
        return CurrencySymbols.TryGetValue(code, out var symbol)
            ? $"{symbol}{n} {code}"
            : $"{n} {code}";
    }
}