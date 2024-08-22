namespace SpreadCheetah.SourceGeneration.Internal;

/// <summary>
/// Instructs the SpreadCheetah source generator to use specified cell mapper to create a cell.
/// </summary>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Property)]
public sealed class CellValueMapperAttribute : Attribute
{
    /// <summary>
    /// Type of mapper to create a cell. Type must be a class and inherit from <see cref="ICellValueMapper{T}"/>
    /// </summary>
    public Type CellValueMapperType { get; set; } = null!;
}