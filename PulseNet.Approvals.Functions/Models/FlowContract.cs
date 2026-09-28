namespace PulseNet.Approvals.Functions.Models;

/// <summary>
/// The wire contract between this app and the Power Automate flow (Option D).
///
/// WHY THESE ARE CONSTANTS AND NOT SETTINGS
///
/// A setting is something one environment does differently from another: a
/// URL, a secret, a timeout, a threshold. These are none of those - they are
/// the shape of the message itself. Changing one here without changing the
/// flow in exactly the same way breaks delivery, so an app setting would only
/// offer a way for the two sides to disagree silently.
///
/// They live in one file so the card builder, the sender and the callback
/// endpoint cannot drift apart, and so the runbook can be checked against a
/// single place.
/// </summary>
public static class FlowContract
{
    // ---- What we POST to the flow (the envelope) ----------------------

    public const string EnvelopeRecipientUpn = "recipientUpn";
    public const string EnvelopeSummary = "summary";
    public const string EnvelopeDocumentNo = "documentNo";
    public const string EnvelopeApprovalEntryNo = "approvalEntryNo";
    public const string EnvelopeEventId = "eventId";
    public const string EnvelopeCorrelationId = "correlationId";
    public const string EnvelopeCardJson = "cardJson";
    public const string EnvelopeExpectsResponse = "expectsResponse";

    // ---- What the card's buttons carry back ---------------------------

    /// <summary>Which button was pressed: <see cref="ActionApprove"/> or <see cref="ActionReject"/>.</summary>
    public const string DataAction = "pnAction";

    /// <summary>The signed action token for that button.</summary>
    public const string DataToken = "pnToken";

    public const string ActionApprove = "approve";
    public const string ActionReject = "reject";

    // ---- How the flow proves it is our flow ---------------------------

    /// <summary>
    /// Header carrying Channels:Teams:FlowCallbackSecret on the callback. The
    /// header NAME is contract; the secret in it is configuration.
    /// </summary>
    public const string SecretHeader = "x-pn-flow-secret";
}
