using SpreadCheetah.Worksheets;
using System.Collections.Immutable;

namespace SpreadCheetah.Test.Helpers;

internal static class SpreadsheetExtensions
{
    public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, Cell cell, RowCollectionType rowType, RowOptions? options = null)
    {
        return spreadsheet.AddRowAsync([cell], rowType, options);
    }

    public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, DataCell cell, RowCollectionType rowType, RowOptions? options = null)
    {
        return spreadsheet.AddRowAsync([cell], rowType, options);
    }

    public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, object obj, RowCollectionType rowType, RowOptions? options = null)
    {
        return spreadsheet.AddRowAsync([obj], rowType, options);
    }

    public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, IList<object> obj, RowCollectionType rowType, RowOptions? options = null) => obj[0] switch
    {
        Cell => spreadsheet.AddRowAsync(obj.Cast<Cell>(), rowType, options),
        DataCell => spreadsheet.AddRowAsync(obj.Cast<DataCell>(), rowType, options),
        StyledCell => spreadsheet.AddRowAsync(obj.Cast<StyledCell>(), rowType, options),
        _ => throw new ArgumentOutOfRangeException(nameof(obj), obj, null)
    };

    public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, IEnumerable<Cell> cells, RowCollectionType rowType, RowOptions? options = null) => options switch
    {
        null => rowType switch
        {
            RowCollectionType.Array => spreadsheet.AddRowAsync(cells.ToArray()),
            RowCollectionType.ReadOnlyMemory => spreadsheet.AddRowAsync(cells.ToArray().AsMemory()),
            RowCollectionType.List => spreadsheet.AddRowAsync(cells.ToList()),
            RowCollectionType.ImmutableList => spreadsheet.AddRowAsync(cells.ToImmutableList()),
            _ => throw new ArgumentOutOfRangeException(nameof(rowType), rowType, null)
        },
        _ => rowType switch
        {
            RowCollectionType.Array => spreadsheet.AddRowAsync(cells.ToArray(), options),
            RowCollectionType.ReadOnlyMemory => spreadsheet.AddRowAsync(cells.ToArray().AsMemory(), options),
            RowCollectionType.List => spreadsheet.AddRowAsync(cells.ToList(), options),
            RowCollectionType.ImmutableList => spreadsheet.AddRowAsync(cells.ToImmutableList(), options),
            _ => throw new ArgumentOutOfRangeException(nameof(rowType), rowType, null)
        }
    };

    public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, IEnumerable<DataCell> cells, RowCollectionType rowType, RowOptions? options = null) => options switch
    {
        null => rowType switch
        {
            RowCollectionType.Array => spreadsheet.AddRowAsync(cells.ToArray()),
            RowCollectionType.ReadOnlyMemory => spreadsheet.AddRowAsync(cells.ToArray().AsMemory()),
            RowCollectionType.List => spreadsheet.AddRowAsync(cells.ToList()),
            RowCollectionType.ImmutableList => spreadsheet.AddRowAsync(cells.ToImmutableList()),
            _ => throw new ArgumentOutOfRangeException(nameof(rowType), rowType, null)
        },
        _ => rowType switch
        {
            RowCollectionType.Array => spreadsheet.AddRowAsync(cells.ToArray(), options),
            RowCollectionType.ReadOnlyMemory => spreadsheet.AddRowAsync(cells.ToArray().AsMemory(), options),
            RowCollectionType.List => spreadsheet.AddRowAsync(cells.ToList(), options),
            RowCollectionType.ImmutableList => spreadsheet.AddRowAsync(cells.ToImmutableList(), options),
            _ => throw new ArgumentOutOfRangeException(nameof(rowType), rowType, null)
        }
    };

    public static ValueTask AddRowAsync(this Spreadsheet spreadsheet, IEnumerable<StyledCell> cells, RowCollectionType rowType, RowOptions? options = null) => options switch
    {
        null => rowType switch
        {
            RowCollectionType.Array => spreadsheet.AddRowAsync(cells.ToArray()),
            RowCollectionType.ReadOnlyMemory => spreadsheet.AddRowAsync(cells.ToArray().AsMemory()),
            RowCollectionType.List => spreadsheet.AddRowAsync(cells.ToList()),
            RowCollectionType.ImmutableList => spreadsheet.AddRowAsync(cells.ToImmutableList()),
            _ => throw new ArgumentOutOfRangeException(nameof(rowType), rowType, null)
        },
        _ => rowType switch
        {
            RowCollectionType.Array => spreadsheet.AddRowAsync(cells.ToArray(), options),
            RowCollectionType.ReadOnlyMemory => spreadsheet.AddRowAsync(cells.ToArray().AsMemory(), options),
            RowCollectionType.List => spreadsheet.AddRowAsync(cells.ToList(), options),
            RowCollectionType.ImmutableList => spreadsheet.AddRowAsync(cells.ToImmutableList(), options),
            _ => throw new ArgumentOutOfRangeException(nameof(rowType), rowType, null)
        }
    };

    public static ValueTask AddHeaderRowAsync(this Spreadsheet spreadsheet, IEnumerable<string> headerNames, RowCollectionType rowType) => rowType switch
    {
        RowCollectionType.Array => spreadsheet.AddHeaderRowAsync(headerNames.ToArray()),
        RowCollectionType.ReadOnlyMemory => spreadsheet.AddHeaderRowAsync(headerNames.ToArray().AsMemory()),
        RowCollectionType.List => spreadsheet.AddHeaderRowAsync(headerNames.ToList()),
        RowCollectionType.ImmutableList => spreadsheet.AddHeaderRowAsync(headerNames.ToImmutableList()),
        _ => throw new ArgumentOutOfRangeException(nameof(rowType), rowType, null)
    };
}
