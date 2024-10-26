using SpreadCheetah.Styling;

namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Instructs the SpreadCheetah source generator to apply a number format style to the cells created for a property.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class CellFormatAttribute : Attribute
{
    /// <summary>
    /// Instructs the SpreadCheetah source generator to apply a custom number format style to the cells created for a property.
    /// </summary>
    public CellFormatAttribute(string customFormat)
    {
    }

    /// <summary>
    /// Instructs the SpreadCheetah source generator to apply a standard number format style to the cells created for a property.
    /// </summary>
    public CellFormatAttribute(StandardNumberFormat format)
    {
    }
}