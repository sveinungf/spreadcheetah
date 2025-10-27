using SpreadCheetah.CellReferences;
using SpreadCheetah.MetadataXml.Attributes;
using SpreadCheetah.Styling;
using SpreadCheetah.Worksheets;

namespace SpreadCheetah.MetadataXml;

internal static class WorksheetStartXml
{
    public static async ValueTask WriteAsync(
        WorksheetOptions? options,
        IReadOnlyDictionary<int, StyleId>? columnStyles,
        SpreadsheetBuffer buffer,
        Stream entryStream,
        CancellationToken token)
    {
        var writer = new WorksheetStartXmlWriter(options, columnStyles, buffer);

        foreach (var success in writer)
        {
            if (!success)
                await buffer.FlushToStreamAsync(entryStream, token).ConfigureAwait(false);
        }

        await buffer.FlushToStreamAsync(entryStream, token).ConfigureAwait(false);
    }
}

file struct WorksheetStartXmlWriter(
    WorksheetOptions? options,
    IReadOnlyDictionary<int, StyleId>? columnStyles,
    SpreadsheetBuffer buffer)
{
    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" """u8 +
        """xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"""u8;

    private static ReadOnlySpan<byte> SheetDataBegin => "<sheetData>"u8;

    private readonly List<(int ColumnIndex, ColumnOptions Options)>? _columns = GetColumns(options?.ColumnOptions);
    private WorksheetColumnXmlPart? _columnXml;
    private Element _next;
    private int _nextIndex;
    private int _previouslyWrittenBytes;
    
    private static List<(int ColumnIndex, ColumnOptions Options)>? GetColumns(
        SortedDictionary<int, ColumnOptions>? dictionary)
    {
        if (dictionary is null)
            return null;

        var columns = new List<(int ColumnIndex, ColumnOptions Options)>(dictionary.Count);

        foreach (var (columnIndex, option) in dictionary)
        {
            if (option is { DefaultStyle: null, Width: null, Hidden: false })
                continue;

            columns.Add((columnIndex, option));
        }

        return columns;
    }

    public readonly WorksheetStartXmlWriter GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        var bytesWritten = 0;
        Current = _next switch
        {
            Element.Header => buffer.TryWrite(Header),
            Element.SheetViewsStart => TryWriteSheetViewsStart(_previouslyWrittenBytes, out bytesWritten),
            Element.SheetViewPane => TryWriteSheetViewPane(_previouslyWrittenBytes, out bytesWritten),
            Element.SheetViewSelection => TryWriteSheetViewSelection(_previouslyWrittenBytes, out bytesWritten),
            Element.SheetViewsEnd => TryWriteSheetViewsEnd(),
            Element.SheetFormatProperties => TryWriteSheetFormatProperties(),
            Element.ColumnsStart => TryWriteColumnsStart(),
            Element.Columns => TryWriteColumns(),
            Element.ColumnsEnd => TryWriteColumnsEnd(),
            _ => buffer.TryWrite(SheetDataBegin)
        };

        if (Current)
        {
            _previouslyWrittenBytes = 0;
            ++_next;
        }
        else
        {
            _previouslyWrittenBytes += bytesWritten;
        }

        return _next < Element.Done;
    }

    private readonly bool TryWriteSheetViewsStart(int bytesToSkip, out int bytesWritten)
    {
        if (options is null or { FrozenColumns: null, FrozenRows: null, ShowGridLines: null })
        {
            bytesWritten = 0;
            return true;
        }

        var showGridLinesAttribute = new BooleanAttribute("showGridLines"u8, options.ShowGridLines);

        return buffer.TryWrite(bytesToSkip, out bytesWritten,
            $"{"<sheetViews><sheetView"u8}" +
            $"{showGridLinesAttribute}" +
            $"{" workbookViewId=\"0\">"u8}");
    }

    private readonly bool TryWriteSheetViewPane(int bytesToSkip, out int bytesWritten)
    {
        if (options is null or { FrozenColumns: null, FrozenRows: null })
        {
            bytesWritten = 0;
            return true;
        }

        var frozenColumns = options.FrozenColumns;
        var frozenRows = options.FrozenRows;

        var xSplitAttribute = new IntAttribute("xSplit"u8, frozenColumns);
        var ySplitAttribute = new IntAttribute("ySplit"u8, frozenRows);

        var activePane = (frozenColumns, frozenRows) switch
        {
            (not null, not null) => "bottomRight"u8,
            (not null, _) => "topRight"u8,
            _ => "bottomLeft"u8
        };

        var activePaneAttribute = new SpanByteAttribute("activePane"u8, activePane);

        var column = (frozenColumns ?? 0) + 1;
        var row = (frozenRows ?? 0) + 1;
        var cellReference = new SimpleSingleCellReference((ushort)column, (uint)row);
        var topLeftCellAttribute = new SimpleSingleCellReferenceAttribute("topLeftCell"u8, cellReference);

        return buffer.TryWrite(bytesToSkip, out bytesWritten,
            $"{"<pane"u8}" +
            $"{xSplitAttribute}" +
            $"{ySplitAttribute}" +
            $"{activePaneAttribute}" +
            $"{topLeftCellAttribute}" +
            $"{" state=\"frozen\"/>"u8}"
        );
    }

    private readonly bool TryWriteSheetViewSelection(int bytesToSkip, out int bytesWritten)
    {
        if (options is null or { FrozenColumns: null, FrozenRows: null })
        {
            bytesWritten = 0;
            return true;
        }

        var activePane = options switch
        {
            { FrozenColumns: not null, FrozenRows: not null } => "bottomRight"u8,
            { FrozenColumns: not null, FrozenRows: _ } => "topRight"u8,
            { FrozenColumns: _, FrozenRows: _ } => "bottomLeft"u8
        };

        var paneAttribute = new SpanByteAttribute("pane"u8, activePane);
        return buffer.TryWrite(bytesToSkip, out bytesWritten,
            $"{"<selection"u8}" +
            $"{paneAttribute}" +
            $"{" />"u8}"
        );
    }

    private readonly bool TryWriteSheetViewsEnd()
    {
        if (options is null or { FrozenColumns: null, FrozenRows: null, ShowGridLines: null })
            return true;

        return buffer.TryWrite("</sheetView></sheetViews>"u8);
    }

    private readonly bool TryWriteSheetFormatProperties()
    {
        if (options?.DefaultColumnWidth is not { } defaultColumnWidth)
            return true;

        return buffer.TryWrite(
            $"{"<sheetFormatPr defaultColWidth=\""u8}" +
            $"{defaultColumnWidth}" +
            $"{"\" defaultRowHeight=\"14.4\"/>"u8}");
    }

    private readonly bool TryWriteColumnsStart()
    {
        return _columns is not { Count: > 0 } || buffer.TryWrite("<cols>"u8);
    }

    private bool TryWriteColumns()
    {
        if (_columns is not { Count: > 0 } columns)
            return true;

        for (; _nextIndex < columns.Count; ++_nextIndex)
        {
            var (columnIndex, columnOptions) = columns[_nextIndex];
            var styleId = columnStyles?.GetValueOrDefault(columnIndex - 1);

            var xml = _columnXml ?? new WorksheetColumnXmlPart(
                buffer: buffer,
                columnIndex: columnIndex,
                options: columnOptions,
                styleId: styleId,
                defaultColumnWidth: options?.DefaultColumnWidth);

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
        return _columns is not { Count: > 0 } || buffer.TryWrite("</cols>"u8);
    }

    private enum Element
    {
        Header,
        SheetViewsStart,
        SheetViewPane,
        SheetViewSelection,
        SheetViewsEnd,
        SheetFormatProperties,
        ColumnsStart,
        Columns,
        ColumnsEnd,
        SheetDataStart,
        Done
    }
}
