namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Instructs the SpreadCheetah source generator how to handle inherited properties.
/// </summary>
/// <param name="strategy">The strategy to generate properties.</param>
[AttributeUsage(AttributeTargets.Class)]
public sealed class InheritedColumnOrderingAttribute(InheritedColumnsOrderingStrategy strategy) : Attribute;