namespace SpreadCheetah.Test.Helpers;

internal static class TestData
{
    private static readonly Type[] CellTypesArray = new[] { typeof(Cell), typeof(DataCell), typeof(StyledCell) };
    private static readonly Type[] StyledCellTypesArray = new[] { typeof(Cell), typeof(StyledCell) };

    public static IEnumerable<object?[]> CellTypes() => CellTypesArray.Select(x => new object[] { x });
    public static IEnumerable<object?[]> StyledCellTypes() => StyledCellTypesArray.Select(x => new object[] { x });

    public static IEnumerable<object?[]> CombineWithCellTypes(params object?[] values)
    {
        return values.SelectMany(_ => CellTypesArray, (value, type) => new object?[] { value, type });
    }

    public static IEnumerable<object?[]> CombineWithCellTypes(params (object?, object?)[] values)
    {
        return values.SelectMany(_ => CellTypesArray, (value, type) => new object?[] { value.Item1, value.Item2, type });
    }

    public static IEnumerable<object?[]> CombineWithStyledCellTypes(params object?[] values)
    {
        return values.SelectMany(_ => StyledCellTypesArray, (value, type) => new object?[] { value, type });
    }
}
