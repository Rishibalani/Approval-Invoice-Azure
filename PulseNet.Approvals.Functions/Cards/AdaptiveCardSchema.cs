namespace PulseNet.Approvals.Functions.Cards;

/// <summary>
/// Adaptive Card wire constants shared by every card this app emits.
///
/// Deliberately constants, not configuration: these identify the card format
/// itself. Changing one does not point the app somewhere else - it produces
/// cards Teams no longer recognises.
/// </summary>
public static class AdaptiveCardSchema
{
    /// <summary>The $schema value on a card. An identifier, never fetched.</summary>
    public const string SchemaUri = "http://adaptivecards.io/schemas/adaptive-card.json";

    /// <summary>Attachment content type for an Adaptive Card in a Bot Framework / Teams message.</summary>
    public const string ContentType = "application/vnd.microsoft.card.adaptive";

    /// <summary>Card version used by the small outcome and notice cards.</summary>
    public const string Version14 = "1.4";
}
