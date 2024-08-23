namespace SpreadCheetah.SourceGeneration;

/// <summary>
/// Cell value mapper marker.
/// </summary>
public interface ICellValueMapper;

/// <summary>
/// Provide the ability to create a custom cell when using source generator.
/// </summary>
/// <typeparam name="T"></typeparam>
public interface ICellValueMapper<in T> : ICellValueMapper
{
    /// <summary>
    /// Map provided value to a cell.
    /// </summary>
    /// <param name="value">The value to wire in cell.</param>
    /// <returns></returns>
    Cell MapToCell(T value);
}