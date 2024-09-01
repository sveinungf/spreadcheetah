namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Instructs the SpreadCheetah source generator to use specified cell mapper to create a cell.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class CellValueConverterAttribute : Attribute
{
    /// <summary>
    /// Type of mapper to create a cell. Type must be a class and inherit from <see cref="CellValueConverter{T}"/>
    /// </summary>
    public Type CellValueConverterType { get; set; } = null!;
}