using Polyfills;

namespace SpreadCheetah.Test.Helpers;

internal static class TestData
{
    private static readonly CellType[] CellTypeArray = EnumPolyfill.GetValues<CellType>();
    private static readonly CellType[] StyledCellTypeArray = [CellType.StyledCell, CellType.Cell];
    private static readonly CellValueType[] CellValueTypeArray = EnumPolyfill.GetValues<CellValueType>();
    private static readonly RowCollectionType[] RowCollectionTypeArray = EnumPolyfill.GetValues<RowCollectionType>();

    public static IEnumerable<object?[]> StyledCellTypes()
    {
        foreach (var cellType in StyledCellTypeArray)
        {
            foreach (var rowType in RowCollectionTypeArray)
            {
                yield return new object?[] { cellType, rowType };
            }
        }
    }

    public static IEnumerable<object?[]> CombineWithCellTypes<T>(params T?[] values)
    {
        foreach (var value in values)
        {
            foreach (var cellType in CellTypeArray)
            {
                foreach (var rowType in RowCollectionTypeArray)
                {
                    yield return new object?[] { value, cellType, rowType };
                }
            }
        }
    }

    public static IEnumerable<object?[]> CombineTuplesWithCellTypes<TFirst, TSecond>(params (TFirst?, TSecond?)[] values)
    {
        foreach (var (firstValue, secondValue) in values)
        {
            foreach (var cellType in CellTypeArray)
            {
                foreach (var rowType in RowCollectionTypeArray)
                {
                    yield return new object?[] { firstValue, secondValue, cellType, rowType };
                }
            }
        }
    }

    public static IEnumerable<object?[]> CombineWithStyledCellTypes<T>(params T?[] values)
    {
        foreach (var value in values)
        {
            foreach (var cellType in StyledCellTypeArray)
            {
                foreach (var rowType in RowCollectionTypeArray)
                {
                    yield return new object?[] { value, cellType, rowType };
                }
            }
        }
    }

    public static IEnumerable<object?[]> CombineWithValueTypes<T>(params T?[] values)
    {
        foreach (var value in values)
        {
            foreach (var valueType in CellValueTypeArray)
            {
                foreach (var rowType in RowCollectionTypeArray)
                {
                    yield return new object?[] { value, valueType, rowType, true };
                    yield return new object?[] { value, valueType, rowType, false };
                }
            }
        }
    }

    public static IEnumerable<object?[]> StyledCellAndValueTypes()
    {
        foreach (var valueType in CellValueTypeArray)
        {
            foreach (var cellType in StyledCellTypeArray)
            {
                foreach (var rowType in RowCollectionTypeArray)
                {
                    yield return new object?[] { valueType, true, cellType, rowType };
                    yield return new object?[] { valueType, false, cellType, rowType };
                }
            }
        }
    }
}
