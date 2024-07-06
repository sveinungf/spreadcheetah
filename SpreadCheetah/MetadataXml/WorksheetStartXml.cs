using SpreadCheetah.Helpers;
using SpreadCheetah.Worksheets;
using System.IO;

namespace SpreadCheetah.MetadataXml;

internal struct WorksheetStartXml
{
    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" """u8 +
        """xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"""u8;

    private static ReadOnlySpan<byte> SheetDataBegin => "<sheetData>"u8;

    private readonly WorksheetOptions? _options;
    private readonly List<KeyValuePair<int, ColumnOptions>>? _columns;
    private readonly SpreadsheetBuffer _buffer;
    private Element _next;
    private int _nextIndex;
    private bool _anyColumnWritten;

    private WorksheetStartXml(WorksheetOptions? options, SpreadsheetBuffer buffer)
    {
        _options = options;
        _columns = options?.ColumnOptions is { } columns ? [.. columns] : null;
        _buffer = buffer;
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
        var destination = _buffer.GetSpan();

        (var current, var bytesWritten) = _next switch
        {
            Element.Header => Header.MyCopyTo(destination),
            Element.SheetViews => TryWriteSheetViews(destination),
            Element.Columns => TryWriteColumns(destination),
            Element.SheetDataBegin => SheetDataBegin.MyCopyTo(destination),
            _ => (true, 0)
        };

        Current = current;
        if (current)
        {
            _buffer.Advance(bytesWritten);
            ++_next;
        }
        

        return _next < Element.Done;
    }

    private readonly (bool, int) TryWriteSheetViews(Span<byte> span)
    {
        var options = _options;
        if (options?.FrozenColumns is null && options?.FrozenRows is null)
            return (true, 0);

        var written = 0;

        if (!"<sheetViews><sheetView workbookViewId=\"0\"><pane "u8.TryCopyTo(span, ref written)) return (false, 0);
        if (!TryWritexSplit(span, ref written)) return (false, 0);
        if (!TryWriteySplit(span, ref written)) return (false, 0);
        if (!"topLeftCell=\""u8.TryCopyTo(span, ref written)) return (false, 0);

        var column = (options.FrozenColumns ?? 0) + 1;
        var row = (options.FrozenRows ?? 0) + 1;
        if (!SpanHelper.TryWriteCellReference(column, (uint)row, span, ref written)) return (false, 0);

        if (!"\" activePane=\""u8.TryCopyTo(span, ref written)) return (false, 0);

        var activePane = options switch
        {
            { FrozenColumns: not null, FrozenRows: not null } => "bottomRight"u8,
            { FrozenColumns: not null } => "topRight"u8,
            _ => "bottomLeft"u8
        };

        if (!activePane.TryCopyTo(span, ref written)) return (false, 0);
        if (!"\" state=\"frozen\"/>"u8.TryCopyTo(span, ref written)) return (false, 0);

        var bottomRight = """<selection pane="bottomRight"/>"""u8;
        if (options is { FrozenColumns: not null, FrozenRows: not null } && !bottomRight.TryCopyTo(span, ref written))
            return (false, 0);

        var bottomLeft = """<selection pane="bottomLeft"/>"""u8;
        if (options.FrozenRows is not null && !bottomLeft.TryCopyTo(span, ref written))
            return (false, 0);

        var topRight = """<selection pane="topRight"/>"""u8;
        if (options.FrozenColumns is not null && !topRight.TryCopyTo(span, ref written))
            return (false, 0);

        if (!"</sheetView></sheetViews>"u8.TryCopyTo(span, ref written)) return (false, 0);

        return (true, written);
    }

    private readonly bool TryWritexSplit(Span<byte> bytes, ref int bytesWritten)
    {
        if (_options?.FrozenColumns is not { } frozenColumns)
            return true;

        if (!"xSplit=\""u8.TryCopyTo(bytes, ref bytesWritten)) return false;
        if (!SpanHelper.TryWrite(frozenColumns, bytes, ref bytesWritten)) return false;
        return "\" "u8.TryCopyTo(bytes, ref bytesWritten);
    }

    private readonly bool TryWriteySplit(Span<byte> bytes, ref int bytesWritten)
    {
        if (_options?.FrozenRows is not { } frozenRows)
            return true;

        if (!"ySplit=\""u8.TryCopyTo(bytes, ref bytesWritten)) return false;
        if (!SpanHelper.TryWrite(frozenRows, bytes, ref bytesWritten)) return false;
        return "\" "u8.TryCopyTo(bytes, ref bytesWritten);
    }

    private (bool, int) TryWriteColumns(Span<byte> span)
    {
        if (_columns is not { } columns) return (true, 0);

        var bytesWritten = 0;
        var anyColumnWritten = _anyColumnWritten;

        for (; _nextIndex < columns.Count; ++_nextIndex)
        {
            var (columnIndex, options) = columns[_nextIndex];
            if (options.Width is null && !options.Hidden)
                continue;

            var written = 0;

            if (!anyColumnWritten)
            {
                if (!"<cols>"u8.TryCopyTo(span, ref written)) return (false, bytesWritten);
                anyColumnWritten = true;
            }

            if (!"<col min=\""u8.TryCopyTo(span, ref written)) return (false, bytesWritten);
            if (!SpanHelper.TryWrite(columnIndex, span, ref written)) return (false, bytesWritten);
            if (!"\" max=\""u8.TryCopyTo(span, ref written)) return (false, bytesWritten);
            if (!SpanHelper.TryWrite(columnIndex, span, ref written)) return (false, bytesWritten);
            if (!"\""u8.TryCopyTo(span, ref written)) return (false, bytesWritten);
            if (options.Width.HasValue)
            {
                if (!" width=\""u8.TryCopyTo(span, ref written)) return (false, bytesWritten);
                if (!SpanHelper.TryWrite(options.Width.GetValueOrDefault(), span, ref written)) return (false, bytesWritten);
                if (!"\" customWidth=\"1\""u8.TryCopyTo(span, ref written)) return (false, bytesWritten);
            }

            if (options.Hidden && !" hidden=\"1\""u8.TryCopyTo(span, ref written)) return (false, bytesWritten);
            if (!" />"u8.TryCopyTo(span, ref written)) return (false, bytesWritten);

            _anyColumnWritten = anyColumnWritten;
            bytesWritten += written;
        }

        if (!anyColumnWritten)
            return (true, bytesWritten);

        return ("</cols>"u8.TryCopyTo(span, ref bytesWritten), bytesWritten);
    }

    private enum Element
    {
        Header,
        SheetViews,
        Columns,
        SheetDataBegin,
        Done
    }
}
