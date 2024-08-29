namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Provide the ability to create a custom cell when using source generator.
/// </summary>
/// <typeparam name="T"></typeparam>
#pragma warning disable S1694
public abstract class CellValueConverter<T>
#pragma warning restore S1694
{
    /// <summary>
    /// Map provided value to a cell.
    /// </summary>
    /// <param name="value">The value to wire in cell.</param>
    /// <returns></returns>
    public abstract DataCell ConvertToCell(T value);
}