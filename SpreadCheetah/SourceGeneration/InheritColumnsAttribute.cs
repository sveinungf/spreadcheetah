namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Instructs the SpreadCheetah source generator to create columns from properties on the base class.
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
public sealed class InheritColumnsAttribute : Attribute
{
    /// <summary>
    /// Specify the default order of inherited columns. Defaults to <see cref="InheritedColumnsOrder.InheritedColumnsFirst"/>.
    /// </summary>
    public InheritedColumnsOrder DefaultColumnOrder { get; set; }
}