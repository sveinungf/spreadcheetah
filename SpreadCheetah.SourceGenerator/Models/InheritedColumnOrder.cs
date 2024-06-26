namespace SpreadCheetah.SourceGenerator.Models;

/// <summary>
/// Specifies the strategy in which the columns will be generated.
/// </summary>
public enum InheritedColumnOrder
{
    /// <summary>
    /// Column generation will start with the properties of the base class.
    /// </summary>
    InheritedColumnsFirst,

    /// <summary>
    /// Column generation will start with the properties of the derived class.
    /// </summary>
    InheritedColumnsLast
}