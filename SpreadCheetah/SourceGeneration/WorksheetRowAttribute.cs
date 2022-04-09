namespace SpreadCheetah.SourceGeneration;

[AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
public sealed class WorksheetRowAttribute : Attribute
{
    public WorksheetRowAttribute(Type type)
    {
    }
}
