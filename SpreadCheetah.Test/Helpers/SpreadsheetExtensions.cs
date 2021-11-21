using SpreadCheetah.Worksheets;

namespace SpreadCheetah.Test.Helpers;

internal static class SpreadsheetExtensions
{
    public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, Cell cell, RowOptions? options = null) => spreadsheet.AddRowAsync(new[] { cell }, options);

    public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, DataCell cell, RowOptions? options = null) => spreadsheet.AddRowAsync(new[] { cell }, options);

    public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, StyledCell cell, RowOptions? options = null) => spreadsheet.AddRowAsync(new[] { cell }, options);

    public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, object obj, RowOptions? options = null) => obj switch
    {
        Cell cell => spreadsheet.AddRowAsync(cell, options),
        DataCell dataCell => spreadsheet.AddRowAsync(dataCell, options),
        StyledCell styledCell => spreadsheet.AddRowAsync(styledCell, options),
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
