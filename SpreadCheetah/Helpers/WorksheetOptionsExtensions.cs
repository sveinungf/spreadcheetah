using SpreadCheetah.Worksheets;
using SpreadCheetah.Worksheets.Internal;

namespace SpreadCheetah.Helpers;

internal static class WorksheetOptionsExtensions
{
    public static List<WorksheetDimensionRun>? GetColumnWidthRuns(this WorksheetOptions worksheetOptions)
    {
        if (worksheetOptions.ColumnOptions is not { } columnOptions)
            return null;

        var runs = new List<WorksheetDimensionRun>();
        WorksheetDimensionRun? currentRun = null;

        foreach (var (columnNumber, options) in columnOptions)
        {
            // TODO: Hidden means width = 0?

            if (options.Width is not { } width)
                continue;

            var number = (uint)columnNumber;

            if (currentRun?.TryContinueWith(number, width) is true)
                continue;

            currentRun = new WorksheetDimensionRun(number, width);
            runs.Add(currentRun);
        }

        return runs;
    }
}
