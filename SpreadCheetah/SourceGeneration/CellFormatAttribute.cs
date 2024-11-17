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
        CustomFormat = customFormat;
    }

    /// <summary>
    /// Instructs the SpreadCheetah source generator to apply a standard number format style to the cells created for a property.
    /// </summary>
    public CellFormatAttribute(StandardNumberFormat format)
    {
        Format = format;
    }


    /// <summary>
    /// <para>When constructed with <see cref="CellFormatAttribute(string)"/>, returns the constructor's <c>customFormat</c> argument value.</para>
    /// <para>When constructed with <see cref="CellFormatAttribute(StandardNumberFormat)"/>, returns null.</para>
    /// </summary>
    public string? CustomFormat { get; }

    /// <summary>
    /// <para>When constructed with <see cref="CellFormatAttribute(StandardNumberFormat)"/>, returns the constructor's <c>format</c> argument value.</para>
    /// <para>When constructed with <see cref="CellFormatAttribute(string)"/>, returns null.</para>
    /// </summary>
    public StandardNumberFormat? Format { get; }
}