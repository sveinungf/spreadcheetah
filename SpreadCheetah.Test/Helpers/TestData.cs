using Polyfills;

namespace SpreadCheetah.Test.Helpers;

internal static class TestData
{
    private static readonly CellType[] StyledCellTypeArray = [CellType.StyledCell, CellType.Cell];
    private static readonly RowCollectionType[] RowCollectionTypeArray = EnumPolyfill.GetValues<RowCollectionType>();

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
}
