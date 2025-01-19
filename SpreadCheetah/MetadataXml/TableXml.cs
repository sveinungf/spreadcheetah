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
        WorksheetTableInfo worksheetTableInfo,
        FileCounter fileCounter,
        CancellationToken token)
    {
        // TODO: Counter here is not correct if there are multiple tables
        var entryName = StringHelper.Invariant($"xl/tables/table{fileCounter.TotalTables}.xml");
        var writer = new TableXmlWriter(
            tableId: fileCounter.TotalTables,
            worksheetTableInfo: worksheetTableInfo,
            buffer: buffer);

        return zipArchiveManager.WriteAsync(writer, entryName, buffer, token);
    }
}

file struct TableXmlWriter(
    int tableId,
    WorksheetTableInfo worksheetTableInfo,
    SpreadsheetBuffer buffer)
    : IXmlWriter<TableXmlWriter>
{
    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="""u8 + "\""u8;

    private Element _next;
    private int _nextIndex;

    private readonly HashSet<string> _uniqueColumnNames = new(StringComparer.OrdinalIgnoreCase);
    private readonly ImmutableTable Table => worksheetTableInfo.Table;
    public readonly TableXmlWriter GetEnumerator() => this;
    public bool Current { get; private set; }

    private readonly int ColumnCount => worksheetTableInfo.ActualNumberOfColumns;

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => TryWriteHeader(),
            Element.Name => buffer.TryWrite($"{Table.Name}"), // TODO: Can it be longer than minimum buffer length? Maybe it should be disallowed?
            Element.NameEnd => buffer.TryWrite("\" displayName=\""u8),
            Element.DisplayName => buffer.TryWrite($"{Table.Name}"), // TODO: Can it be longer than minimum buffer length? Maybe it should be disallowed?
            Element.DisplayNameEnd => buffer.TryWrite("\" ref=\""u8),
            Element.Reference => TryWriteTableReference(),
            Element.ReferenceEnd => TryWriteReferenceEnd(),
            Element.AutoFilter => TryWriteAutoFilter(),
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
        var toRow = lastDataRow + (Table.HasTotalRow ? 1u : 0u);
        return buffer.TryWrite($"{FromCell()}{":"u8}{ToCell(toRow)}");
    }

    private readonly SimpleSingleCellReference FromCell()
    {
        var firstColumn = worksheetTableInfo.FirstColumn;
        return new SimpleSingleCellReference(firstColumn, worksheetTableInfo.FirstRow);
    }

    private readonly SimpleSingleCellReference ToCell(uint toRow)
    {
        var firstColumn = worksheetTableInfo.FirstColumn;
        var endColumn = (ushort)(firstColumn + ColumnCount - 1);
        return new SimpleSingleCellReference(endColumn, toRow);
    }

    private readonly bool TryWriteReferenceEnd()
    {
        return buffer.TryWrite(
            $"{"\" headerRowCount=\""u8}" +
            $"{worksheetTableInfo.HasHeaderRow}" +
            $"{"\" totalsRowCount=\""u8}" +
            $"{Table.HasTotalRow}" +
            $"{"\">"u8}");
    }

    private readonly bool TryWriteAutoFilter()
    {
        if (!worksheetTableInfo.HasHeaderRow)
            return true;

        var lastDataRow = worksheetTableInfo.LastDataRow ?? 0;
        return buffer.TryWrite(
            $"{"<autoFilter ref=\""u8}" +
            $"{FromCell()}{":"u8}{ToCell(lastDataRow)}" +
            $"{"\"/>"u8}");
    }

    private readonly bool TryWriteTableColumnsStart()
    {
        Debug.Assert(ColumnCount > 0);
        return buffer.TryWrite(
            $"{"<tableColumns count=\""u8}" +
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
            if (name is null || string.IsNullOrEmpty(name))
                name = TableNameGenerator.GenerateUniqueTableColumnName(_uniqueColumnNames);

            // TODO: Need to make changes to the name?
            // TODO: Need unique names here. Consider validating in AddHeaderRow

            var (label, function) = Table.ColumnOptions is { } columns && columns.TryGetValue(_nextIndex + 1, out var options)
                ? (options.TotalRowLabel, options.TotalRowFunction)
                : (null, null);

            if (!TryWriteTableColumn(name, label, function))
                return false;

            _uniqueColumnNames.Add(name);
        }

        for (; _nextIndex < ColumnCount; ++_nextIndex)
        {
            var name = TableNameGenerator.GenerateUniqueTableColumnName(_uniqueColumnNames);

            if (!TryWriteTableColumn(name, null, null))
                return false;

            _uniqueColumnNames.Add(name);
        }

        _nextIndex = 0;
        return true;
    }

    private readonly bool TryWriteTableColumn(string name,
        string? totalRowLabel,
        TableTotalRowFunction? totalRowFunction)
    {
        var labelAttribute = totalRowLabel is not null ? "\" totalsRowLabel=\""u8 : [];
        var functionAttribute = totalRowFunction is (null or TableTotalRowFunction.None) ? [] : "\" totalsRowFunction=\""u8;
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
        TableTotalRowFunction.Count => "count"u8,
        TableTotalRowFunction.CountNumbers => "countNums"u8,
        TableTotalRowFunction.Maximum => "max"u8,
        TableTotalRowFunction.Minimum => "min"u8,
        TableTotalRowFunction.StandardDeviation => "stdDev"u8,
        TableTotalRowFunction.Sum => "sum"u8,
        TableTotalRowFunction.Variance => "var"u8,
        _ => []
    };

    private readonly bool TryWriteTableStyleName()
    {
        var style = Table.Style;
        if (style is TableStyle.None)
            return true;

        Debug.Assert(style is > TableStyle.None and <= TableStyle.Dark11);

        var styleCategory = style switch
        {
            >= TableStyle.Dark1 => "Dark"u8,
            >= TableStyle.Medium1 => "Medium"u8,
            _ => "Light"u8
        };

        var styleNumber = style switch
        {
            >= TableStyle.Dark1 => style - TableStyle.Dark1 + 1,
            >= TableStyle.Medium1 => style - TableStyle.Medium1 + 1,
            _ => (int)style
        };

        return buffer.TryWrite($"{"name=\"TableStyle"u8}{styleCategory}{styleNumber}{"\" "u8}");
    }

    private readonly bool TryWriteTableStyleInfo()
    {
        return buffer.TryWrite(
            $"{"showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\""u8}" +
            $"{Table.BandedRows}" +
            $"{"\" showColumnStripes=\""u8}" +
            $"{Table.BandedColumns}" +
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
    AutoFilter,
    TableColumnsStart,
    TableColumns,
    TableStyleInfoStart,
    TableStyleName,
    TableStyleInfo,
    Footer,
    Done
}