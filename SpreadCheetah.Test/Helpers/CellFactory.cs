using SpreadCheetah.Styling;

namespace SpreadCheetah.Test.Helpers;

internal static class CellFactory
{
    public static object CreateWithoutValue(Type type)
    {
        if (type == typeof(Cell))
            return new Cell();
        if (type == typeof(DataCell))
            return new DataCell();
        if (type == typeof(StyledCell))
            return new StyledCell();

        throw new ArgumentOutOfRangeException(nameof(type), type, null);
    }

    public static object Create(Type type, string? value)
    {
        if (type == typeof(Cell))
            return new Cell(value, null);
        if (type == typeof(DataCell))
            return new DataCell(value);
        if (type == typeof(StyledCell))
            return new StyledCell(value, null);

        throw new ArgumentOutOfRangeException(nameof(type), type, null);
    }

    public static object Create(Type type, int? value)
    {
        if (type == typeof(Cell))
            return new Cell(value, null);
        if (type == typeof(DataCell))
            return new DataCell(value);
        if (type == typeof(StyledCell))
            return new StyledCell(value, null);

        throw new ArgumentOutOfRangeException(nameof(type), type, null);
    }

    public static object Create(Type type, long? value)
    {
        if (type == typeof(Cell))
            return new Cell(value, null);
        if (type == typeof(DataCell))
            return new DataCell(value);
        if (type == typeof(StyledCell))
            return new StyledCell(value, null);

        throw new ArgumentOutOfRangeException(nameof(type), type, null);
    }

    public static object Create(Type type, float? value)
    {
        if (type == typeof(Cell))
            return new Cell(value, null);
        if (type == typeof(DataCell))
            return new DataCell(value);
        if (type == typeof(StyledCell))
            return new StyledCell(value, null);

        throw new ArgumentOutOfRangeException(nameof(type), type, null);
    }

    public static object Create(Type type, double? value)
    {
        if (type == typeof(Cell))
            return new Cell(value, null);
        if (type == typeof(DataCell))
            return new DataCell(value);
        if (type == typeof(StyledCell))
            return new StyledCell(value, null);

        throw new ArgumentOutOfRangeException(nameof(type), type, null);
    }

    public static object Create(Type type, decimal? value)
    {
        if (type == typeof(Cell))
            return new Cell(value, null);
        if (type == typeof(DataCell))
            return new DataCell(value);
        if (type == typeof(StyledCell))
            return new StyledCell(value, null);

        throw new ArgumentOutOfRangeException(nameof(type), type, null);
    }

    public static object Create(Type type, bool? value)
    {
        if (type == typeof(Cell))
            return new Cell(value, null);
        if (type == typeof(DataCell))
            return new DataCell(value);
        if (type == typeof(StyledCell))
            return new StyledCell(value, null);

        throw new ArgumentOutOfRangeException(nameof(type), type, null);
    }

    public static object Create(Type type, string? value, StyleId? styleId)
    {
        if (type == typeof(Cell))
            return new Cell(value, styleId);
        if (type == typeof(StyledCell))
            return new StyledCell(value, styleId);

        throw new ArgumentOutOfRangeException(nameof(type), type, null);
    }
}
