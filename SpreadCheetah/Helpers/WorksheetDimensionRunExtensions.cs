using SpreadCheetah.Worksheets.Internal;

namespace SpreadCheetah.Helpers;

internal static class WorksheetDimensionRunExtensions
{
    public static IEnumerable<double> GetColumnWidths(this List<WorksheetDimensionRun>? runs)
    {
        return runs.GetDimensions(SpreadsheetConstants.DefaultColumnWidth, SpreadsheetConstants.MaxNumberOfColumns);
    }

    public static IEnumerable<double> GetRowHeights(this List<WorksheetDimensionRun>? runs)
    {
        return runs.GetDimensions(SpreadsheetConstants.DefaultRowHeight, SpreadsheetConstants.MaxNumberOfRows);
    }

    private static IEnumerable<double> GetDimensions(
        this List<WorksheetDimensionRun>? runs,
        double defaultValue,
        int maxIterations)
    {
        return runs is { Count: > 0 }
            ? runs.GetDimensionsInternal(defaultValue, maxIterations)
            : Enumerable.Repeat(defaultValue, maxIterations);
    }

    private static IEnumerable<double> GetDimensionsInternal(
        this List<WorksheetDimensionRun> runs,
        double defaultValue,
        int maxIterations)
    {
        var runIndex = 0;
        var currentRun = runs[0];

        for (var i = 1u; i <= maxIterations; ++i)
        {
            if (currentRun is null)
            {
                yield return defaultValue;
                continue;
            }

            if (i > currentRun.LastIndex)
            {
                runIndex++;
                currentRun = runs.ElementAtOrDefault(runIndex);
            }

            yield return currentRun?.ContainsIndex(i) is true
                ? currentRun.Dimension
                : defaultValue;
        }
    }
}
