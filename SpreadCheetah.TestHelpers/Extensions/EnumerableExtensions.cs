using SpreadCheetah.TestHelpers.Interfaces;

namespace SpreadCheetah.TestHelpers.Extensions;

public static class EnumerableExtensions
{
    public static IEnumerable<string?> StringValues(this IEnumerable<IWorksheetCell> cells)
    {
        return cells.Select(x => x.StringValue);
    }
}
