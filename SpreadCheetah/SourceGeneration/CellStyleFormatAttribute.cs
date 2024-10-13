using SpreadCheetah.Styling;

namespace SpreadCheetah.SourceGeneration;

[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class CellStyleFormatAttribute : Attribute
{
    public CellStyleFormatAttribute(string customFormat)
    {
    }

    public CellStyleFormatAttribute(StandardNumberFormat format)
    {
    }
}