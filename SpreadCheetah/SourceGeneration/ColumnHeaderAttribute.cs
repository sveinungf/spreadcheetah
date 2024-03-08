namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Instructs the SpreadCheetah source generator about what header name to use for a column.
/// The source generator creates columns from properties, and using the attribute on a property will set the header name for the resulting column.
/// Header names are written to a worksheet with <see cref="Spreadsheet.AddHeaderRowAsync"/>.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ColumnHeaderAttribute : Attribute
{
    public ColumnHeaderAttribute(string name)
    {
    }

    public ColumnHeaderAttribute(Type type, string propertyName)
    {
    }
}
