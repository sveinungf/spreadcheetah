namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Instructs the SpreadCheetah source generator to use specified cell mapper to create a cell.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class CellValueConverterAttribute(Type CellValueConverterType) : Attribute
{

}