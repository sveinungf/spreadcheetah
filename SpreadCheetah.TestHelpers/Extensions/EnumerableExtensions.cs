using SpreadCheetah.TestHelpers.Assertions;

namespace SpreadCheetah.TestHelpers.Extensions;

public static class EnumerableExtensions
{
    public static IEnumerable<string?> StringValues(this IEnumerable<ISpreadsheetAssertCell> cells)
    {
        return cells.Select(x => x.StringValue);
    }
}
