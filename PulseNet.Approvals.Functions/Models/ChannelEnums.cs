namespace PulseNet.Approvals.Functions.Models;

/// <summary>
/// HOW a message physically reaches the approver.
///
/// The point of this enum is that upgrading a channel is a configuration
/// change, not a code change. Teams starts on WorkflowWebhook because that
/// needs nothing but a URL, and moves to Bot when the registration clears.
/// Nothing in the dispatch pipeline knows the difference.
/// </summary>
public enum ChannelDeliveryMode
{
    /// <summary>Master off switch for this channel.</summary>
    Disabled = 0,

    /// <summary>Teams, via a Workflows incoming webhook URL. Zero registration.</summary>
    WorkflowWebhook = 1,

    /// <summary>Plain HTML email with link buttons. Zero registration.</summary>
    PlainEmail = 2,

    /// <summary>Teams, via Azure Bot Service. Needs app registration + Teams admin upload.</summary>
    Bot = 3,

    /// <summary>Outlook Actionable Message. Needs provider registration + Entra token validation.</summary>
    ActionableMessage = 4,

    /// <summary>WhatsApp utility template via Meta Cloud API. Needs template approval.</summary>
    WhatsAppTemplate = 5
}

/// <summary>
/// WHAT the buttons do, which is a separate question from how the message is sent.
///
/// A Teams card sent by webhook can carry Link buttons; the same card sent by a
/// bot can carry Native ones. Keeping these orthogonal is what lets a single
/// card builder serve every tier.
/// </summary>
public enum ChannelActionMode
{
    /// <summary>
    /// No action buttons. Title, amount, counterparty and a deep link into
    /// Business Central. Used above a threshold, or when a gate has fired.
    /// </summary>
    NotifyOnly = 0,

    /// <summary>
    /// Action.OpenUrl in Teams, an anchor tag in email, a URL button on a
    /// WhatsApp template. All three point at the same Function endpoint
    /// carrying a signed token. Works on every channel with no registration.
    /// </summary>
    Link = 1,

    /// <summary>
    /// Channel-native interaction: Action.Execute, Action.Http, quick-reply.
    /// Better UX, but each one needs its own registration.
    /// </summary>
    Native = 2
}

public enum ApprovalAction
{
    Approve = 0,
    Reject = 1
}

public enum ApprovalChannel
{
    None = 0,
    Teams = 1,
    Outlook = 2,
    WhatsApp = 3
}
