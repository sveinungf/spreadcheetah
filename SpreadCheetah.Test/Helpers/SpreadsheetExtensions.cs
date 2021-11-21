namespace SpreadCheetah.Test.Helpers;

internal static class SpreadsheetExtensions
{
    public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, Cell cell) => spreadsheet.AddRowAsync(new[] { cell });

    public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, DataCell cell) => spreadsheet.AddRowAsync(new[] { cell });

    public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, StyledCell cell) => spreadsheet.AddRowAsync(new[] { cell });

    public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, object obj) => obj switch
    {
        Cell cell => spreadsheet.AddRowAsync(cell),
        DataCell dataCell => spreadsheet.AddRowAsync(dataCell),
        StyledCell styledCell => spreadsheet.AddRowAsync(styledCell),
        _ => throw new ArgumentOutOfRangeException(nameof(obj), obj, null)
    };

    public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, IList<object> obj) => obj[0] switch
    {
        Cell => spreadsheet.AddRowAsync(obj.Cast<Cell>().ToList()),
        DataCell => spreadsheet.AddRowAsync(obj.Cast<DataCell>().ToList()),
        StyledCell => spreadsheet.AddRowAsync(obj.Cast<StyledCell>().ToList()),
        _ => throw new ArgumentOutOfRangeException(nameof(obj), obj, null)
    };
}
