using SpreadCheetah.Styling;

namespace SpreadCheetah.Test.Helpers;

internal static class CellFactory
{
    public static object CreateWithoutValue(CellType type) => type switch
    {
        CellType.Cell => new Cell(),
        CellType.DataCell => new DataCell(),
        CellType.StyledCell => new StyledCell(),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(CellType type, int? value) => type switch
    {
        CellType.Cell => new Cell(value, null),
        CellType.DataCell => new DataCell(value),
        CellType.StyledCell => new StyledCell(value, null),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(CellType type, long? value) => type switch
    {
        CellType.Cell => new Cell(value, null),
        CellType.DataCell => new DataCell(value),
        CellType.StyledCell => new StyledCell(value, null),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(CellType type, float? value) => type switch
    {
        CellType.Cell => new Cell(value, null),
        CellType.DataCell => new DataCell(value),
        CellType.StyledCell => new StyledCell(value, null),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(CellType type, double? value) => type switch
    {
        CellType.Cell => new Cell(value, null),
        CellType.DataCell => new DataCell(value),
        CellType.StyledCell => new StyledCell(value, null),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(CellType type, decimal? value) => type switch
    {
        CellType.Cell => new Cell(value, null),
        CellType.DataCell => new DataCell(value),
        CellType.StyledCell => new StyledCell(value, null),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(CellType type, DateTime? value) => type switch
    {
        CellType.Cell => new Cell(value, null),
        CellType.DataCell => new DataCell(value),
        CellType.StyledCell => new StyledCell(value, null),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(StyledCellType type, DateTime? value, StyleId? styleId) => type switch
    {
        StyledCellType.Cell => new Cell(value, styleId),
        StyledCellType.StyledCell => new StyledCell(value, styleId),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(CellType type, bool? value) => type switch
    {
        CellType.Cell => new Cell(value, null),
        CellType.DataCell => new DataCell(value),
        CellType.StyledCell => new StyledCell(value, null),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(CellType type, string? value) => type switch
    {
        CellType.Cell => new Cell(value, null),
        CellType.DataCell => new DataCell(value),
        CellType.StyledCell => new StyledCell(value, null),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(CellType type, ReadOnlyMemory<char> value) => type switch
    {
        CellType.Cell => new Cell(value),
        CellType.DataCell => new DataCell(value),
        CellType.StyledCell => new StyledCell(value, null),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(CellType type, string? value, StyleId? styleId) => type switch
    {
        CellType.Cell => new Cell(value, styleId),
        CellType.StyledCell => new StyledCell(value, styleId),
        CellType.DataCell => new DataCell(value),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(StyledCellType type, string? value, StyleId? styleId) => type switch
    {
        StyledCellType.Cell => new Cell(value, styleId),
        StyledCellType.StyledCell => new StyledCell(value, styleId),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
    };

    public static object Create(StyledCellType styledCellType, CellValueType valueType, bool isNull, StyleId? styleId, out object? value)
    {
        var cellType = styledCellType.ToCellType();
        return Create(cellType, valueType, isNull, styleId, out value);
    }

    public static object Create(CellType cellType, CellValueType valueType, bool isNull, StyleId? styleId, out object? value) => valueType switch
    {
        CellValueType.BoolFalse => CreateForBool(cellType, isNull ? null : false, styleId, out value),
        CellValueType.BoolTrue => CreateForBool(cellType, isNull ? null : true, styleId, out value),
        CellValueType.DateTime => CreateForDateTime(cellType, isNull, styleId, out value),
        CellValueType.Decimal => CreateForDecimal(cellType, isNull, styleId, out value),
        CellValueType.Double => CreateForDouble(cellType, isNull, styleId, out value),
        CellValueType.Float => CreateForFloat(cellType, isNull, styleId, out value),
        CellValueType.Int => CreateForInt(cellType, isNull, styleId, out value),
        CellValueType.Long => CreateForLong(cellType, isNull, styleId, out value),
        CellValueType.String => CreateForString(cellType, isNull, styleId, out value),
        _ => throw new ArgumentOutOfRangeException(nameof(valueType), valueType, null)
    };

    public static Cell Create(Formula formula, CellValueType valueType, bool isNull, StyleId? styleId, out object? value) => valueType switch
    {
        CellValueType.BoolFalse => CreateForBool(formula, isNull ? null : false, styleId, out value),
        CellValueType.BoolTrue => CreateForBool(formula, isNull ? null : true, styleId, out value),
        CellValueType.DateTime => CreateForDateTime(formula, isNull, styleId, out value),
        CellValueType.Decimal => CreateForDecimal(formula, isNull, styleId, out value),
        CellValueType.Double => CreateForDouble(formula, isNull, styleId, out value),
        CellValueType.Float => CreateForFloat(formula, isNull, styleId, out value),
        CellValueType.Int => CreateForInt(formula, isNull, styleId, out value),
        CellValueType.Long => CreateForLong(formula, isNull, styleId, out value),
        CellValueType.String => CreateForString(formula, isNull, styleId, out value),
        _ => throw new ArgumentOutOfRangeException(nameof(valueType), valueType, null)
    };

    private static object CreateForBool(CellType cellType, bool? actualValue, StyleId? styleId, out object? value)
    {
        value = actualValue;

        return cellType switch
        {
            CellType.DataCell => new DataCell(actualValue),
            CellType.StyledCell => new StyledCell(actualValue, styleId),
            CellType.Cell => new Cell(actualValue, styleId),
            _ => throw new ArgumentOutOfRangeException(nameof(cellType), cellType, null)
        };
    }

    private static Cell CreateForBool(Formula formula, bool? actualValue, StyleId? styleId, out object? value)
    {
        value = actualValue;
        return new Cell(formula, actualValue, styleId);
    }

    private static object CreateForDateTime(CellType cellType, bool isNull, StyleId? styleId, out object? value)
    {
        DateTime? actualValue = !isNull ? new DateTime(2021, 2, 3, 0, 0, 0, DateTimeKind.Unspecified) : null;
        value = actualValue;

        return cellType switch
        {
            CellType.DataCell => new DataCell(actualValue),
            CellType.StyledCell => new StyledCell(actualValue, styleId),
            CellType.Cell => new Cell(actualValue, styleId),
            _ => throw new ArgumentOutOfRangeException(nameof(cellType), cellType, null)
        };
    }

    private static Cell CreateForDateTime(Formula formula, bool isNull, StyleId? styleId, out object? value)
    {
        DateTime? actualValue = !isNull ? new DateTime(2021, 2, 3, 0, 0, 0, DateTimeKind.Unspecified) : null;
        value = actualValue;
        return new Cell(formula, actualValue, styleId);
    }

    private static object CreateForDecimal(CellType cellType, bool isNull, StyleId? styleId, out object? value)
    {
        decimal? actualValue = !isNull ? 9.8m : null;
        value = actualValue;

        return cellType switch
        {
            CellType.DataCell => new DataCell(actualValue),
            CellType.StyledCell => new StyledCell(actualValue, styleId),
            CellType.Cell => new Cell(actualValue, styleId),
            _ => throw new ArgumentOutOfRangeException(nameof(cellType), cellType, null)
        };
    }

    private static Cell CreateForDecimal(Formula formula, bool isNull, StyleId? styleId, out object? value)
    {
        decimal? actualValue = !isNull ? 9.8m : null;
        value = actualValue;
        return new Cell(formula, actualValue, styleId);
    }

    private static object CreateForDouble(CellType cellType, bool isNull, StyleId? styleId, out object? value)
    {
        double? actualValue = !isNull ? 8.7 : null;
        value = actualValue;

        return cellType switch
        {
            CellType.DataCell => new DataCell(actualValue),
            CellType.StyledCell => new StyledCell(actualValue, styleId),
            CellType.Cell => new Cell(actualValue, styleId),
            _ => throw new ArgumentOutOfRangeException(nameof(cellType), cellType, null)
        };
    }

    private static Cell CreateForDouble(Formula formula, bool isNull, StyleId? styleId, out object? value)
    {
        double? actualValue = !isNull ? 8.7 : null;
        value = actualValue;
        return new Cell(formula, actualValue, styleId);
    }

    private static object CreateForFloat(CellType cellType, bool isNull, StyleId? styleId, out object? value)
    {
        float? actualValue = !isNull ? 7.6f : null;
        value = actualValue;

        return cellType switch
        {
            CellType.DataCell => new DataCell(actualValue),
            CellType.StyledCell => new StyledCell(actualValue, styleId),
            CellType.Cell => new Cell(actualValue, styleId),
            _ => throw new ArgumentOutOfRangeException(nameof(cellType), cellType, null)
        };
    }

    private static Cell CreateForFloat(Formula formula, bool isNull, StyleId? styleId, out object? value)
    {
        float? actualValue = !isNull ? 7.6f : null;
        value = actualValue;
        return new Cell(formula, actualValue, styleId);
    }

    private static object CreateForInt(CellType cellType, bool isNull, StyleId? styleId, out object? value)
    {
        int? actualValue = !isNull ? 12 : null;
        value = actualValue;

        return cellType switch
        {
            CellType.DataCell => new DataCell(actualValue),
            CellType.StyledCell => new StyledCell(actualValue, styleId),
            CellType.Cell => new Cell(actualValue, styleId),
            _ => throw new ArgumentOutOfRangeException(nameof(cellType), cellType, null)
        };
    }

    private static Cell CreateForInt(Formula formula, bool isNull, StyleId? styleId, out object? value)
    {
        int? actualValue = !isNull ? 12 : null;
        value = actualValue;
        return new Cell(formula, actualValue, styleId);
    }

    private static object CreateForLong(CellType cellType, bool isNull, StyleId? styleId, out object? value)
    {
        long? actualValue = !isNull ? 23L : null;
        value = actualValue;

        return cellType switch
        {
            CellType.DataCell => new DataCell(actualValue),
            CellType.StyledCell => new StyledCell(actualValue, styleId),
            CellType.Cell => new Cell(actualValue, styleId),
            _ => throw new ArgumentOutOfRangeException(nameof(cellType), cellType, null)
        };
    }

    private static Cell CreateForLong(Formula formula, bool isNull, StyleId? styleId, out object? value)
    {
        long? actualValue = !isNull ? 23L : null;
        value = actualValue;
        return new Cell(formula, actualValue, styleId);
    }

    private static object CreateForString(CellType cellType, bool isNull, StyleId? styleId, out object? value)
    {
        string? actualValue = !isNull ? "abc" : null;
        value = actualValue;

        return cellType switch
        {
            CellType.DataCell => new DataCell(actualValue),
            CellType.StyledCell => new StyledCell(actualValue, styleId),
            CellType.Cell => new Cell(actualValue, styleId),
            _ => throw new ArgumentOutOfRangeException(nameof(cellType), cellType, null)
        };
    }

    private static Cell CreateForString(Formula formula, bool isNull, StyleId? styleId, out object? value)
    {
        string? actualValue = !isNull ? "abc" : null;
        value = actualValue;
        return new Cell(formula, actualValue, styleId);
    }

    private static CellType ToCellType(this StyledCellType fromType) => fromType switch
    {
        StyledCellType.Cell => CellType.Cell,
        StyledCellType.StyledCell => CellType.StyledCell,
        _ => throw new ArgumentOutOfRangeException(nameof(fromType), fromType, null)
    };
}
