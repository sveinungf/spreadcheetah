namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Converts a value to a <see cref="DataCell"/>. Used in conjunction with <see cref="CellValueConverterAttribute"/>.
/// </summary>
public abstract class CellValueConverter<T>
{
    /// <summary>
    /// Converts a value to a <see cref="DataCell"/>.
    /// </summary>
    public abstract DataCell ConvertToCell(T value);
}