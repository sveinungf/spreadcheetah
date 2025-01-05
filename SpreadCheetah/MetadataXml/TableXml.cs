using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;
using SpreadCheetah.Tables;
using SpreadCheetah.Tables.Internal;
using System.Diagnostics;

namespace SpreadCheetah.MetadataXml;

internal static class TableXml
{
    public static ValueTask WriteAsync(
        ZipArchiveManager zipArchiveManager,
        SpreadsheetBuffer buffer,
        Table table,
        WorksheetTableInfo worksheetTableInfo,
        FileCounter fileCounter,
        CancellationToken token)
    {
        var entryName = StringHelper.Invariant($"xl/tables/table{fileCounter.TotalTables}.xml");
        var writer = new TableXmlWriter(
            table: table,
            tableId: fileCounter.TotalTables,
            worksheetTableInfo: worksheetTableInfo,
            buffer: buffer);

        return zipArchiveManager.WriteAsync(writer, entryName, buffer, token);
    }
}

file struct TableXmlWriter(
    Table table,
    int tableId,
    WorksheetTableInfo worksheetTableInfo,
    SpreadsheetBuffer buffer)
    : IXmlWriter<TableXmlWriter>
{
    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="""u8 + "\""u8;

    private readonly bool _hasTotalRow = table.HasTotalRow();
    private Element _next;
    private int _nextIndex;

    public readonly TableXmlWriter GetEnumerator() => this;
    public bool Current { get; private set; }

    // TODO: Can ColumnCount end up being "0"?
    private readonly int ColumnCount => table.NumberOfColumns ?? worksheetTableInfo.HeaderNames.Length;

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => TryWriteHeader(),
            Element.Name => buffer.TryWrite($"{table.Name}"), // TODO: Can it be longer than minimum buffer length? Maybe it should be disallowed?
            Element.NameEnd => buffer.TryWrite("\" displayName=\""u8),
            Element.DisplayName => buffer.TryWrite($"{table.Name}"), // TODO: Can it be longer than minimum buffer length? Maybe it should be disallowed?
            Element.DisplayNameEnd => buffer.TryWrite("\" ref=\""u8),
            Element.Reference => TryWriteTableReference(),
            Element.ReferenceEnd => TryWriteReferenceEnd(),
            Element.AutoFilterReference => TryWriteAutoFilterReference(),
            Element.TableColumnsStart => TryWriteTableColumnsStart(),
            Element.TableColumns => TryWriteTableColumns(),
            Element.TableStyleInfoStart => buffer.TryWrite("</tableColumns><tableStyleInfo "u8),
            Element.TableStyleName => TryWriteTableStyleName(),
            Element.TableStyleInfo => TryWriteTableStyleInfo(),
            _ => buffer.TryWrite("</table>"u8)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteHeader()
    {
        return buffer.TryWrite($"{Header}{tableId}{"\" name=\""u8}");
    }

    private readonly bool TryWriteTableReference()
    {
        var lastDataRow = worksheetTableInfo.LastDataRow ?? 0;
        Debug.Assert(lastDataRow > 0);
        var toRow = lastDataRow + (_hasTotalRow ? 1u : 0u);
        return TryWriteCellRangeReference(toRow);
    }

    private readonly bool TryWriteAutoFilterReference()
    {
        var lastDataRow = worksheetTableInfo.LastDataRow ?? 0;
        Debug.Assert(lastDataRow > 0);
        return TryWriteCellRangeReference(lastDataRow);
    }

    private readonly bool TryWriteCellRangeReference(uint toRow)
    {
        var firstColumn = worksheetTableInfo.FirstColumn;
        var endColumn = (ushort)(firstColumn + ColumnCount);
        var fromCell = new SimpleSingleCellReference(firstColumn, worksheetTableInfo.FirstRow);
        var toCell = new SimpleSingleCellReference(endColumn, toRow);

        return buffer.TryWrite($"{fromCell}:{toCell}");
    }

    private readonly bool TryWriteReferenceEnd()
    {
        // TODO: AutoFilter by default for all columns. Can later add options to change this behavior.
        return buffer.TryWrite(
            $"{"\" totalsRowShown=\""u8}" +
            $"{_hasTotalRow}" +
            $"{"\"><autoFilter ref=\""u8}");
    }

    private readonly bool TryWriteTableColumnsStart()
    {
        return buffer.TryWrite(
            $"{"\"/><tableColumns count=\""u8}" +
            $"{ColumnCount}" +
            $"{"\">"u8}");
    }

    private bool TryWriteTableColumns()
    {
        var headerNames = worksheetTableInfo.HeaderNames;
        var length = Math.Min(headerNames.Length, ColumnCount);

        for (; _nextIndex < length; ++_nextIndex)
        {
            var name = headerNames[_nextIndex];
            // TODO: Need to make changes to the name?
            // TODO: What to do when name is null or empty?
            // TODO: Need unique names here? If so, might have to validate in AddHeaderRow (depends on wether Excel allows duplicate header names or not)

            var (label, function) = table.ColumnOptions is { } columns && columns.TryGetValue(_nextIndex + 1, out var options)
                ? (options.TotalRowLabel, options.TotalRowFunction)
                : (null, null);

            if (!TryWriteTableColumn(name, label, function))
                return false;
        }

        for (; _nextIndex < ColumnCount; ++_nextIndex)
        {
            var name = "Column1"; // TODO: Generate a unique name, e.g. "ColumnX"

            if (!TryWriteTableColumn(name, null, null))
                return false;
        }

        _nextIndex = 0;
        return true;
    }

    private readonly bool TryWriteTableColumn(string name,
        string? totalRowLabel,
        TableTotalRowFunction? totalRowFunction)
    {
        var labelAttribute = totalRowLabel is not null ? "\" totalsRowLabel=\""u8 : [];
        var functionAttribute = totalRowFunction is not null ? "\" totalsRowFunction=\""u8 : [];
        var functionAttributeValue = GetFunctionAttributeValue(totalRowFunction);

        return buffer.TryWrite(
            $"{"<tableColumn id=\""u8}" +
            $"{_nextIndex + 1}" +
            $"{"\" name=\""u8}" +
            $"{name}" +
            $"{labelAttribute}" +
            $"{totalRowLabel}" +
            $"{functionAttribute}" +
            $"{functionAttributeValue}" +
            $"{"\"/>"u8}");
    }

    private static ReadOnlySpan<byte> GetFunctionAttributeValue(TableTotalRowFunction? function) => function switch
    {
        TableTotalRowFunction.Average => "average"u8,
        TableTotalRowFunction.Count => "count"u8, // TODO: Verify
        TableTotalRowFunction.CountNumbers => "countNums"u8, // TODO: Verify
        TableTotalRowFunction.Maximum => "max"u8, // TODO: Verify
        TableTotalRowFunction.Minimum => "min"u8, // TODO: Verify
        TableTotalRowFunction.StandardDeviation => "stdDev"u8, // TODO: Verify
        TableTotalRowFunction.Sum => "sum"u8,
        TableTotalRowFunction.Variance => "var"u8, // TODO: Verify
        _ => []
    };

    private readonly bool TryWriteTableStyleName()
    {
        if (table.Style is TableStyle.None)
            return true;

        Debug.Assert(table.Style is > TableStyle.None and <= TableStyle.Dark11);

        var styleCategory = table.Style switch
        {
            >= TableStyle.Dark1 => "Dark"u8,
            >= TableStyle.Medium1 => "Medium"u8,
            _ => "Light"u8
        };

        var styleNumber = table.Style switch
        {
            >= TableStyle.Dark1 => table.Style - TableStyle.Dark1 + 1,
            >= TableStyle.Medium1 => table.Style - TableStyle.Medium1 + 1,
            _ => (int)table.Style
        };

        return buffer.TryWrite($"{"TableStyle"u8}{styleCategory}{styleNumber}");
    }

    private readonly bool TryWriteTableStyleInfo()
    {
        return buffer.TryWrite(
            $"{"\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\""u8}" +
            $"{table.BandedRows}" +
            $"{"\" showColumnStripes=\""u8}" +
            $"{table.BandedColumns}" +
            $"{"\"/>"u8}");
    }
}

file enum Element
{
    Header,
    Name,
    NameEnd,
    DisplayName,
    DisplayNameEnd,
    Reference,
    ReferenceEnd,
    AutoFilterReference,
    TableColumnsStart,
    TableColumns,
    TableStyleInfoStart,
    TableStyleName,
    TableStyleInfo,
    Footer,
    Done
}