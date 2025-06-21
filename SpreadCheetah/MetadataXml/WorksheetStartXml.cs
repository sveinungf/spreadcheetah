using SpreadCheetah.CellReferences;
using SpreadCheetah.Worksheets;

namespace SpreadCheetah.MetadataXml;

internal struct WorksheetStartXml
{
    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" """u8 +
        """xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"""u8;

    private static ReadOnlySpan<byte> SheetDataBegin => "<sheetData>"u8;

    private readonly WorksheetOptions? _options;
    private readonly List<(int ColumnIndex, ColumnOptions Options)>? _columns;
    private readonly SpreadsheetBuffer _buffer;
    private WorksheetColumnXmlPart? _columnXml;
    private Element _next;
    private int _nextIndex;

    private WorksheetStartXml(WorksheetOptions? options, SpreadsheetBuffer buffer)
    {
        _options = options;
        _buffer = buffer;

        if (options?.ColumnOptions is not { } dictionary)
            return;

        _columns = new List<(int ColumnIndex, ColumnOptions Options)>(dictionary.Count);

        foreach (var (columnIndex, option) in dictionary)
        {
            if (option.Width is not null || option.Hidden)
                _columns.Add((columnIndex, option));
        }
    }

    public static async ValueTask WriteAsync(
        WorksheetOptions? options,
        SpreadsheetBuffer buffer,
        Stream stream,
        CancellationToken token)
    {
        var writer = new WorksheetStartXml(options, buffer);

        foreach (var success in writer)
        {
            if (!success)
                await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }
    }

    public readonly WorksheetStartXml GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => _buffer.TryWrite(Header),
            Element.SheetViewsStart => TryWriteSheetViewsStart(),
            Element.ColumnSplit => TryWriteColumnSplit(),
            Element.RowSplit => TryWriteRowSplit(),
            Element.SheetViewsEnd => TryWriteSheetViewsEnd(),
            Element.ColumnsStart => TryWriteColumnsStart(),
            Element.Columns => TryWriteColumns(),
            Element.ColumnsEnd => TryWriteColumnsEnd(),
            _ => _buffer.TryWrite(SheetDataBegin)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteSheetViewsStart()
    {
        return _options is null or { FrozenColumns: null, FrozenRows: null }
            || _buffer.TryWrite("<sheetViews><sheetView workbookViewId=\"0\"><pane "u8);
    }

    private readonly bool TryWriteColumnSplit()
    {
        if (_options?.FrozenColumns is not { } frozenColumns)
            return true;

        return _buffer.TryWrite(
            $"{"xSplit=\""u8}" +
            $"{frozenColumns}" +
            $"{"\" "u8}");
    }

    private readonly bool TryWriteRowSplit()
    {
        if (_options?.FrozenRows is not { } frozenRows)
            return true;

        return _buffer.TryWrite(
            $"{"ySplit=\""u8}" +
            $"{frozenRows}" +
            $"{"\" "u8}");
    }

    private readonly bool TryWriteSheetViewsEnd()
    {
        var frozenColumns = _options?.FrozenColumns;
        var frozenRows = _options?.FrozenRows;
        if (frozenColumns is null && frozenRows is null)
            return true;

        var activePane = (frozenColumns, frozenRows) switch
        {
            (not null, not null) => "bottomRight"u8,
            (not null, _) => "topRight"u8,
            _ => "bottomLeft"u8
        };

        var column = (frozenColumns ?? 0) + 1;
        var row = (frozenRows ?? 0) + 1;
        var cellReference = new SimpleSingleCellReference((ushort)column, (uint)row);

        return _buffer.TryWrite(
            $"{"topLeftCell=\""u8}" +
            $"{cellReference}" +
            $"{"\" activePane=\""u8}" +
            $"{activePane}" +
            $"{"\" state=\"frozen\"/><selection pane=\""u8}" +
            $"{activePane}" +
            $"{"\"/></sheetView></sheetViews>"u8}");
    }

    private readonly bool TryWriteColumnsStart()
    {
        return _columns is not { Count: > 0 } || _buffer.TryWrite("<cols>"u8);
    }

    private bool TryWriteColumns()
    {
        if (_columns is not { Count: > 0 } columns)
            return true;

        for (; _nextIndex < columns.Count; ++_nextIndex)
        {
            var (columnIndex, options) = columns[_nextIndex];

            var xml = _columnXml ??
                new WorksheetColumnXmlPart(_buffer, columnIndex, options);

            if (!xml.TryWrite())
            {
                _columnXml = xml;
                return false;
            }

            _columnXml = null;
        }

        _nextIndex = 0;
        return true;
    }

    private readonly bool TryWriteColumnsEnd()
    {
        return _columns is not { Count: > 0 } || _buffer.TryWrite("</cols>"u8);
    }

    private enum Element
    {
        Header,
        SheetViewsStart,
        ColumnSplit,
        RowSplit,
        SheetViewsEnd,
        ColumnsStart,
        Columns,
        ColumnsEnd,
        SheetDataStart,
        Done
    }
}
