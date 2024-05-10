namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Instructs the SpreadCheetah source generator how to handle inherited properties.
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
public sealed class InheritColumnsAttribute : Attribute
{
    /// <summary>
    /// Set default order for inherited properties.
    /// </summary>
    public InheritColumnsAttribute()
    {
        InheritedColumnOrder = InheritedColumnOrder.InheritedColumnsFirst;
    }

    /// <summary>
    /// Specify the order of inherited properties.
    /// </summary>
    public InheritedColumnOrder InheritedColumnOrder { get; set; }
}