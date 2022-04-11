namespace SpreadCheetah.SourceGeneration;

[AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
public sealed class WorksheetRowGenerationOptionsAttribute : Attribute
{
    public bool SuppressWarnings { get; set; }
}
