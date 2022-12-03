using SpreadCheetah.Test.Helpers.Backporting;

namespace SpreadCheetah.Test.Helpers;

internal static class TestData
{
    private static readonly CellType[] CellTypeEnum = EnumHelper.GetValues<CellType>();
    private static readonly CellType[] StyledCellTypeEnum = new[] { CellType.StyledCell, CellType.Cell };
    private static readonly CellValueType[] CellValueTypeEnum = EnumHelper.GetValues<CellValueType>();
    private static readonly RowCollectionType[] RowCollectionTypeEnum = EnumHelper.GetValues<RowCollectionType>();

    private static readonly (CellType CellType, RowCollectionType RowType)[] CellTypesValues = CellTypeEnum.SelectMany(_ => RowCollectionTypeEnum, (c, r) => (c, r)).ToArray();
    private static readonly (CellType CellType, RowCollectionType RowType)[] StyledCellTypesValues = StyledCellTypeEnum.SelectMany(_ => RowCollectionTypeEnum, (c, r) => (c, r)).ToArray();

    public static IEnumerable<object?[]> CellTypes() => CellTypesValues.Select(x => new object?[] { x.CellType, x.RowType });
    public static IEnumerable<object?[]> StyledCellTypes() => StyledCellTypesValues.Select(x => new object?[] { x.CellType, x.RowType });

    public static IEnumerable<object?[]> CombineWithCellTypes<T>(params T?[] values)
    {
        return values.SelectMany(_ => CellTypesValues, (value, element) => new object?[] { value, element.CellType, element.RowType });
    }

    public static IEnumerable<object?[]> CombineWithCellTypes<TFirst, TSecond>(params (TFirst?, TSecond?)[] values)
    {
        return values.SelectMany(_ => CellTypesValues, (value, element) => new object?[] { value.Item1, value.Item2, element.CellType, element.RowType });
    }

    // TODO: Replace
    public static IEnumerable<object?[]> CombineWithStyledCellTypes<T>(params T?[] values)
    {
        return values.SelectMany(_ => StyledCellTypeEnum, (value, type) => new object?[] { value, type });
    }

    // TODO: Replace
    public static IEnumerable<object?[]> CombineWithValueTypes(params object?[] values)
    {
        return values.SelectMany(_ => ValueTypes(), (value, type) => new object?[] { value, type[0], type[1] });
    }

    // TODO: Replace
    public static IEnumerable<object?[]> ValueTypes()
    {
        foreach (var valueType in CellValueTypeEnum)
        {
            yield return new object?[] { valueType, true };
            yield return new object?[] { valueType, false };
        }
    }

    // TODO: Replace
    public static IEnumerable<object?[]> StyledCellAndValueTypes()
    {
        foreach (var valueType in CellValueTypeEnum)
        {
            foreach (var cellType in StyledCellTypeEnum)
            {
                yield return new object?[] { valueType, true, cellType };
                yield return new object?[] { valueType, false, cellType };
            }
        }
    }
}
