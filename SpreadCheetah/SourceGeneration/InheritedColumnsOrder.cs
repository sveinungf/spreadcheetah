namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Specifies the order for <see cref="InheritColumnsAttribute"/>.
/// </summary>
public enum InheritedColumnsOrder
{
    /// <summary>
    /// Columns from the base class will be ordered before columns from the derived class.
    /// </summary>
    InheritedColumnsFirst,

    /// <summary>
    /// Columns from the derived class will be ordered before columns from the base class.
    /// </summary>
    InheritedColumnsLast
}