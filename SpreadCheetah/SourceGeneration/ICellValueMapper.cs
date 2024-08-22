namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Provide the ability to create custom cell when using source generator.
/// </summary>
/// <typeparam name="T"></typeparam>
public interface ICellValueMapper<in T>
{
    /// <summary>
    /// Map provided value to a cell.
    /// </summary>
    /// <param name="value">The value to wire in cell.</param>
    /// <param name="styleName">The style name to apply.</param>
    /// <returns></returns>
    Cell MapToCell(T value, string? styleName);
}