namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Instructs the SpreadCheetah source generator to automatically infer the header names for columns.
/// The source generator will try to use properties on the resource type based on the property names from the class having this attribute.
/// E.g. for property <c>MyProperty</c> the source generator will try to use property <c>MyProperty</c> on the resource type for the column header.
/// Additionally, a prefix and/or a suffix can be used, if the resource type is following a naming convention.
/// E.g. if <c>Prefix</c> is set to <c>"Header_"</c>, then the source generator will try to use property <c>Header_MyProperty</c> instead for the column header.
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
public sealed class InferColumnHeadersAttribute(Type type) : Attribute
{
    /// <summary>
    /// The resource type.
    /// </summary>
    public Type? Type { get; } = type;

    /// <summary>
    /// Optional prefix for the resource type property names.
    /// </summary>
    public string? Prefix { get; set; }

    /// <summary>
    /// Optional suffix for the resource type property names.
    /// </summary>
    public string? Suffix { get; set; }
}
