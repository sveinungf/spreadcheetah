namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Instructs the SpreadCheetah source generator to truncate the cell values for a property to the specified length.
/// The attribute is currently only supported on <see langword="string"/> properties.
/// The specified length must be greater than 0.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class CellValueTruncateAttribute(int length) : Attribute
{
    /// <summary>
    /// Returns the constructor's <c>length</c> argument value.
    /// </summary>
    public int Length { get; } = length;
}
