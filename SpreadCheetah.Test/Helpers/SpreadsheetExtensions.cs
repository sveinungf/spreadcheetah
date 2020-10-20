using System.Threading.Tasks;

namespace SpreadCheetah.Test.Helpers
{
    internal static class SpreadsheetExtensions
    {
        public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, Cell cell) => spreadsheet.AddRowAsync(new[] { cell });

        public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, StyledCell cell) => spreadsheet.AddRowAsync(new[] { cell });
    }
}
