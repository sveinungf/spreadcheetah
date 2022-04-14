namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Options for the SpreadCheetah source generator.
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
public sealed class WorksheetRowGenerationOptionsAttribute : Attribute
{
    /// <summary>
    /// Specifies whether to suppress warnings from the source generator.
    /// </summary>
    public bool SuppressWarnings { get; set; }
}
