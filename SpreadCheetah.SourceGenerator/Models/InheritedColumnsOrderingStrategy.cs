namespace SpreadCheetah.SourceGenerator.Models;

/// <summary>
/// Specifies the strategy in which the columns will be generated.
/// </summary>
public enum InheritedColumnsOrderingStrategy
{
    /// <summary>
    /// Properties from base class will be ignored and not generated.
    /// </summary>
    IgnoreInheritedProperties,
    /// <summary>
    /// Column generation will start with the properties of the parent class.
    /// </summary>
    StartFromInheritedProperties,
    /// <summary>
    /// Column generation will start with the properties of the child class.
    /// </summary>
    StartFromClassProperties
}