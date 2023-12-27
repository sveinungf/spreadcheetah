using SpreadCheetah.Helpers;
using SpreadCheetah.Worksheets;

namespace SpreadCheetah.MetadataXml;

internal struct WorksheetStartXml : IXmlWriter
{
    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" """u8 +
        """xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"""u8;

    private static ReadOnlySpan<byte> SheetDataBegin => "<sheetData>"u8;

    private readonly WorksheetOptions? _options;
    private readonly List<KeyValuePair<int, ColumnOptions>>? _columns;
    private Element _next;
    private int _nextIndex;
    private bool _anyColumnWritten;

    public WorksheetStartXml(WorksheetOptions? options)
    {
        _options = options;
        _columns = options?.ColumnOptions.ToList();
    }

    public bool TryWrite(Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;

        if (_next == Element.Header && !Advance(Header.TryCopyTo(bytes, ref bytesWritten))) return false;
        if (_next == Element.SheetViews && !Advance(TryWriteSheetViews(bytes, ref bytesWritten))) return false;
        if (_next == Element.Columns && !Advance(TryWriteColumns(bytes, ref bytesWritten))) return false;
        if (_next == Element.SheetDataBegin && !Advance(SheetDataBegin.TryCopyTo(bytes, ref bytesWritten))) return false;

        return true;
    }

    private bool Advance(bool success)
    {
        if (success)
            ++_next;

        return success;
    }

    private readonly bool TryWriteSheetViews(Span<byte> bytes, ref int bytesWritten)
    {
        var options = _options;
        if (options?.FrozenColumns is null && options?.FrozenRows is null)
            return true;

        var span = bytes.Slice(bytesWritten);
        var written = 0;

        if (!"<sheetViews><sheetView workbookViewId=\"0\"><pane "u8.TryCopyTo(span, ref written)) return false;
        if (!TryWritexSplit(span, ref written)) return false;
        if (!TryWriteySplit(span, ref written)) return false;
        if (!"topLeftCell=\""u8.TryCopyTo(span, ref written)) return false;

        var column = (options.FrozenColumns ?? 0) + 1;
        var row = (options.FrozenRows ?? 0) + 1;
        if (!SpanHelper.TryWriteCellReference(column, (uint)row, span, ref written)) return false;

        if (!"\" activePane=\""u8.TryCopyTo(span, ref written)) return false;

        var activePane = options switch
        {
            { FrozenColumns: not null, FrozenRows: not null } => "bottomRight"u8,
            { FrozenColumns: not null } => "topRight"u8,
            _ => "bottomLeft"u8
        };

        if (!activePane.TryCopyTo(span, ref written)) return false;
        if (!"\" state=\"frozen\"/>"u8.TryCopyTo(span, ref written)) return false;

        var bottomRight = """<selection pane="bottomRight"/>"""u8;
        if (options is { FrozenColumns: not null, FrozenRows: not null } && !bottomRight.TryCopyTo(span, ref written))
            return false;

        var bottomLeft = """<selection pane="bottomLeft"/>"""u8;
        if (options.FrozenRows is not null && !bottomLeft.TryCopyTo(span, ref written))
            return false;

        var topRight = """<selection pane="topRight"/>"""u8;
        if (options.FrozenColumns is not null && !topRight.TryCopyTo(span, ref written))
            return false;

        if (!"</sheetView></sheetViews>"u8.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
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

    private bool TryWriteColumns(Span<byte> bytes, ref int bytesWritten)
    {
        if (_columns is not { } columns) return true;

        var anyColumnWritten = _anyColumnWritten;

        for (; _nextIndex < columns.Count; ++_nextIndex)
        {
            var (columnIndex, options) = columns[_nextIndex];
            if (options.Width is not { } width)
                continue;

            var span = bytes.Slice(bytesWritten);
            var written = 0;

            if (!anyColumnWritten)
            {
                if (!"<cols>"u8.TryCopyTo(span, ref written)) return false;
                anyColumnWritten = true;
            }

            if (!"<col min=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(columnIndex, span, ref written)) return false;
            if (!"\" max=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(columnIndex, span, ref written)) return false;
            if (!"\" width=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(width, span, ref written)) return false;
            if (!"\" customWidth=\"1\" />"u8.TryCopyTo(span, ref written)) return false;

            _anyColumnWritten = anyColumnWritten;
            bytesWritten += written;
        }

        return !anyColumnWritten || "</cols>"u8.TryCopyTo(bytes, ref bytesWritten);
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
