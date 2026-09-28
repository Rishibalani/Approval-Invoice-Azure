using Microsoft.Extensions.Configuration;

namespace PulseNet.Approvals.Functions.Options;

/// <summary>
/// The architectural switch between Teams delivery mechanisms.
///
/// WHY A TOP-LEVEL SETTING RATHER THAN Channels:Teams:DeliveryMode
///
/// Both paths stay fully configured at the same time: the Azure Bot settings,
/// the bot sender and the bot's own endpoint are untouched and ready. This one
/// setting decides which of them carries an approval today, so switching back
/// is a setting change and a restart, not a deployment.
///
/// The app-setting name is exactly USE_POWER_AUTOMATE_CARDS, with no section
/// prefix. Program.cs reads and parses it explicitly and configures this
/// instance rather than binding the configuration root - see the comment there.
/// The [ConfigurationKeyName] below records the key name and keeps the class
/// bindable should that ever be wanted.
/// </summary>
public sealed class FeatureFlagOptions
{
    /// <summary>
    /// true  - Option D: our Adaptive Card is handed to a Power Automate flow,
    ///         which posts it as the Workflows (Flow) bot and calls the flow
    ///         callback endpoint with the press.
    /// false - Option A: the custom Teams bot sends the card and answers the
    ///         press on /api/messages, exactly as before.
    ///
    /// There is deliberately no default. A missing value stops the host at
    /// startup rather than silently picking an architecture.
    /// </summary>
    [ConfigurationKeyName("USE_POWER_AUTOMATE_CARDS")]
    public bool UsePowerAutomateCards { get; set; }
}
