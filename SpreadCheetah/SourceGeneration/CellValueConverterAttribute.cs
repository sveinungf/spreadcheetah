namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Instructs the SpreadCheetah source generator to convert property values to cells using the specified type.
/// The specified converter type must derive from <see cref="CellValueConverter{T}"/>,
/// and its type parameter must match the property type.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class CellValueConverterAttribute(Type converterType) : Attribute;