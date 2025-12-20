using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;
using SpreadCheetah.Tables;
using SpreadCheetah.Tables.Internal;
using System.Diagnostics;

namespace SpreadCheetah.MetadataXml.Tables;

internal static class TableXml
{
    public static async ValueTask WriteTablesAsync(
        this ZipArchiveManager zipArchiveManager,
        List<WorksheetTableInfo> tables,
        SpreadsheetBuffer buffer,
        FileCounter? fileCounter,
        CancellationToken token)
    {
        var fileStartIndex = fileCounter?.CurrentWorksheetTableFileStartIndex ?? 0;
        Debug.Assert(fileStartIndex > 0);

        foreach (var (i, table) in tables.Index())
        {
            var tableId = fileStartIndex + i;
            var entryName = StringHelper.Invariant($"xl/tables/table{tableId}.xml");
            var writer = new TableXmlWriter(
                tableId: tableId,
                worksheetTableInfo: table,
                buffer: buffer);

            await zipArchiveManager.WriteAsync(writer, entryName, buffer, token).ConfigureAwait(false);
        }
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

    private TableColumnXmlPart? _columnXml;
    private Element _next;
    private int _nextIndex;
    private int _currentNameIndex;

    private readonly ImmutableTable Table => worksheetTableInfo.Table;
    public readonly TableXmlWriter GetEnumerator() => this;
    public bool Current { get; private set; }

    private readonly int ColumnCount => worksheetTableInfo.ActualNumberOfColumns;

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => TryWriteHeader(),
            Element.Name => TryWriteName(),
            Element.NameEnd => buffer.TryWrite("\" displayName=\""u8),
            Element.DisplayName => TryWriteName(),
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

    private bool TryWriteName()
    {
        if (buffer.WriteLongStringXmlEncoded(Table.Name, ref _currentNameIndex))
        {
            _currentNameIndex = 0;
            return true;
        }

        return false;
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
        var columnOptions = Table.ColumnOptions;

        for (; _nextIndex < ColumnCount; ++_nextIndex)
        {
            if (_columnXml is not { } xml)
            {
                var headerName = worksheetTableInfo.GetHeaderName(_nextIndex);
                var options = columnOptions?.GetValueOrDefault(_nextIndex + 1);

                xml = new TableColumnXmlPart(
                    buffer,
                    _nextIndex,
                    XmlUtility.XmlEncode(headerName),
                    XmlUtility.XmlEncode(options?.TotalRowLabel),
                    options?.TotalRowFunction);
            }

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