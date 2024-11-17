namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Instructs the SpreadCheetah source generator about what header name to use for a column.
/// The source generator creates columns from properties, and using the attribute on a property will set the header name for the resulting column.
/// Header names are written to a worksheet with <see cref="Spreadsheet.AddHeaderRowAsync"/>.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ColumnHeaderAttribute : Attribute
{
    /// <summary>
    /// Use the value of <c>name</c> as the header name for the column.
    /// </summary>
    public ColumnHeaderAttribute(string name)
    {
        Name = name;
    }

    /// <summary>
    /// Get the header name from a property. The property must:
    /// <list type="bullet">
    ///   <item><description>Be a <see langword="static"/> property.</description></item>
    ///   <item><description>Have a public getter.</description></item>
    ///   <item><description>Have a return type of <see langword="string"/> (or <see langword="string?"/>).</description></item>
    /// </list>
    /// </summary>
    public ColumnHeaderAttribute(Type type, string propertyName)
    {
        Type = type;
        PropertyName = propertyName;
    }


    /// <summary>
    /// <para>When constructed with <see cref="ColumnHeaderAttribute(string)"/>, returns the constructor's <c>name</c> argument value.</para>
    /// <para>When constructed with <see cref="ColumnHeaderAttribute(Type, string)"/>, returns null.</para>
    /// </summary>
    public string? Name { get; }

    /// <summary>
    /// <para>When constructed with <see cref="ColumnHeaderAttribute(Type, string)"/>, returns the constructor's <c>type</c> argument value.</para>
    /// <para>When constructed with <see cref="ColumnHeaderAttribute(string)"/>, returns null.</para>
    /// </summary>
    public Type? Type { get; }

    /// <summary>
    /// <para>When constructed with <see cref="ColumnHeaderAttribute(Type, string)"/>, returns the constructor's <c>propertyName</c> argument value.</para>
    /// <para>When constructed with <see cref="ColumnHeaderAttribute(string)"/>, returns null.</para>
    /// </summary>
    public string? PropertyName { get; }

    
}
