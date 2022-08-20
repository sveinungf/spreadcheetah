using SpreadCheetah.Styling;

namespace SpreadCheetah.Test.Helpers;

internal static class CellFactory
{
    public static object CreateWithoutValue(Type type) => type switch
    {
        _ when type == typeof(Cell) => new Cell(),
        _ when type == typeof(DataCell) => new DataCell(),
        _ when type == typeof(StyledCell) => new StyledCell(),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(Type type, string? value) => type switch
    {
        _ when type == typeof(Cell) => new Cell(value, null),
        _ when type == typeof(DataCell) => new DataCell(value),
        _ when type == typeof(StyledCell) => new StyledCell(value, null),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(Type type, int? value) => type switch
    {
        _ when type == typeof(Cell) => new Cell(value, null),
        _ when type == typeof(DataCell) => new DataCell(value),
        _ when type == typeof(StyledCell) => new StyledCell(value, null),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(Type type, long? value) => type switch
    {
        _ when type == typeof(Cell) => new Cell(value, null),
        _ when type == typeof(DataCell) => new DataCell(value),
        _ when type == typeof(StyledCell) => new StyledCell(value, null),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(Type type, float? value) => type switch
    {
        _ when type == typeof(Cell) => new Cell(value, null),
        _ when type == typeof(DataCell) => new DataCell(value),
        _ when type == typeof(StyledCell) => new StyledCell(value, null),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(Type type, double? value) => type switch
    {
        _ when type == typeof(Cell) => new Cell(value, null),
        _ when type == typeof(DataCell) => new DataCell(value),
        _ when type == typeof(StyledCell) => new StyledCell(value, null),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(Type type, decimal? value) => type switch
    {
        _ when type == typeof(Cell) => new Cell(value, null),
        _ when type == typeof(DataCell) => new DataCell(value),
        _ when type == typeof(StyledCell) => new StyledCell(value, null),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(Type type, DateTime? value) => type switch
    {
        _ when type == typeof(Cell) => new Cell(value, null),
        _ when type == typeof(DataCell) => new DataCell(value),
        _ when type == typeof(StyledCell) => new StyledCell(value, null),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(Type type, bool? value) => type switch
    {
        _ when type == typeof(Cell) => new Cell(value, null),
        _ when type == typeof(DataCell) => new DataCell(value),
        _ when type == typeof(StyledCell) => new StyledCell(value, null),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(Type type, string? value, StyleId? styleId) => type switch
    {
        _ when type == typeof(Cell) => new Cell(value, styleId),
        _ when type == typeof(StyledCell) => new StyledCell(value, styleId),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };
}
