namespace SpreadCheetah.Test.Helpers
{
    internal static class SpreadsheetExtensions
    {
        public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, Cell cell) => spreadsheet.AddRowAsync(new[] { cell });

        public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, DataCell cell) => spreadsheet.AddRowAsync(new[] { cell });

        public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, StyledCell cell) => spreadsheet.AddRowAsync(new[] { cell });

        public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, object obj)
        {
            if (obj is Cell cell)
                return spreadsheet.AddRowAsync(cell);
            if (obj is DataCell dataCell)
                return spreadsheet.AddRowAsync(dataCell);
            if (obj is StyledCell styledCell)
                return spreadsheet.AddRowAsync(styledCell);

            throw new ArgumentOutOfRangeException(nameof(obj), obj, null);
        }

        public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, IList<object> obj)
        {
            if (obj[0] is Cell)
                return spreadsheet.AddRowAsync(obj.Cast<Cell>().ToList());
            if (obj[0] is DataCell)
                return spreadsheet.AddRowAsync(obj.Cast<DataCell>().ToList());
            if (obj[0] is StyledCell)
                return spreadsheet.AddRowAsync(obj.Cast<StyledCell>().ToList());

            throw new ArgumentOutOfRangeException(nameof(obj), obj, null);
        }
    }
}
