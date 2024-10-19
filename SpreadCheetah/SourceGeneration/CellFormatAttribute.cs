using SpreadCheetah.Styling;

namespace SpreadCheetah.SourceGeneration;

[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class CellFormatAttribute : Attribute
{
    public CellFormatAttribute(string customFormat)
    {
    }

    public CellFormatAttribute(StandardNumberFormat format)
    {
    }
}