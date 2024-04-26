namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Instructs the SpreadCheetah source generator how to handle inherited properties.
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public sealed class InheritColumnsAttribute : Attribute
{
    /// <summary>
    /// Set default order for inherited properties.
    /// </summary>
    /// <param name="defaultOrder"></param>
    public InheritColumnsAttribute(InheritedColumnOrder defaultOrder = InheritedColumnOrder.InheritedColumnsFirst)
    {

    }
}