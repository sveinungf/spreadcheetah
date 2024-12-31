using SpreadCheetah.CellReferences;
using SpreadCheetah.Tables;
using System.Diagnostics;

namespace SpreadCheetah.MetadataXml;

internal struct TableXml
{
    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="""u8 + "\""u8;

    private readonly SpreadsheetBuffer _buffer;
    private readonly Table _table;
    private readonly ReadOnlyMemory<string?> _headerRowNames; // TODO: Set value. Might be empty. Consider how to handle tables that don't start at colum 'A'.
    private readonly ushort _startColumn; // TODO: Set value
    private readonly uint _startRow; // TODO: Set value
    private readonly uint _endRow; // TODO: Set value
    private readonly int _tableId; // TODO: Table id set from TotalNumberOfTables
    private Element _next;
    private int _nextIndex;

    private TableXml(Table table, int tableId, SpreadsheetBuffer buffer)
    {
        _buffer = buffer;
        _table = table;
        _tableId = tableId;
    }

    public readonly TableXml GetEnumerator() => this;
    public bool Current { get; private set; }

    // TODO: Can ColumnCount end up being "0"?
    private readonly int ColumnCount => _table.NumberOfColumns ?? _headerRowNames.Length;

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => TryWriteHeader(),
            Element.Name => _buffer.TryWrite($"{_table.Name}"), // TODO: Can it be longer than minimum buffer length? Maybe it should be disallowed?
            Element.NameEnd => _buffer.TryWrite("\" displayName=\""u8),
            Element.DisplayName => _buffer.TryWrite($"{_table.Name}"), // TODO: Can it be longer than minimum buffer length? Maybe it should be disallowed?
            Element.DisplayNameEnd => _buffer.TryWrite("\" ref=\""u8),
            Element.Reference => TryWriteCellRangeReference(),
            Element.ReferenceEnd => TryWriteReferenceEnd(),
            Element.AutoFilterReference => TryWriteCellRangeReference(),
            Element.TableColumnsStart => TryWriteTableColumnsStart(),
            Element.TableColumns => TryWriteTableColumns(),
            Element.TableStyleInfoStart => _buffer.TryWrite("</tableColumns><tableStyleInfo "u8),
            Element.TableStyleName => TryWriteTableStyleName(),
            Element.TableStyleInfo => TryWriteTableStyleInfo(),
            _ => _buffer.TryWrite("</table>"u8)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteHeader()
    {
        return _buffer.TryWrite($"{Header}{_tableId}{"\" name=\""u8}");
    }

    private readonly bool TryWriteCellRangeReference()
    {
        var endColumn = (ushort)(_startColumn + ColumnCount);
        var fromCell = new SimpleSingleCellReference(_startColumn, _startRow);
        var toCell = new SimpleSingleCellReference(endColumn, _endRow);

        return _buffer.TryWrite($"{fromCell}:{toCell}");
    }

    private readonly bool TryWriteReferenceEnd()
    {
        var hasTotalRow = _table.ColumnOptions?.Values.Any(x => x.AffectsTotalRow) ?? false;

        // TODO: AutoFilter by default for all columns. Can later add options to change this behavior.
        return _buffer.TryWrite(
            $"{"\" totalsRowShown=\""u8}" +
            $"{hasTotalRow}" +
            $"{"\"><autoFilter ref=\""u8}");
    }

    private readonly bool TryWriteTableColumnsStart()
    {
        return _buffer.TryWrite(
            $"{"\"/><tableColumns count=\""u8}" +
            $"{ColumnCount}" +
            $"{"\">"u8}");
    }

    private bool TryWriteTableColumns()
    {
        var span = _headerRowNames.Span;
        var length = Math.Min(span.Length, ColumnCount);

        for (; _nextIndex < length; ++_nextIndex)
        {
            var name = span[_nextIndex];
            // TODO: Need to make changes to the name?
            // TODO: What to do when name is null or empty?
            // TODO: Need unique names here? If so, might have to validate in AddHeaderRow

            if (!TryWriteTableColumn(name))
                return false;
        }

        for (; _nextIndex < ColumnCount; ++_nextIndex)
        {
            var name = ""; // TODO: Generate a unique name, e.g. "ColumnX"

            if (!TryWriteTableColumn(name))
                return false;
        }

        _nextIndex = 0;
        return true;
    }

    private readonly bool TryWriteTableColumn(string name)
    {
        return _buffer.TryWrite(
            $"{"<tableColumn id=\""u8}" +
            $"{_nextIndex + 1}" +
            $"{"\" name=\""u8}" +
            $"{name}" +
            $"{"\"/>"u8}");
    }

    private readonly bool TryWriteTableStyleName()
    {
        if (_table.Style is TableStyle.None)
            return true;

        Debug.Assert(_table.Style is > TableStyle.None and <= TableStyle.Dark11);

        var styleCategory = _table.Style switch
        {
            >= TableStyle.Dark1 => "Dark"u8,
            >= TableStyle.Medium1 => "Medium"u8,
            _ => "Light"u8
        };

        var styleNumber = _table.Style switch
        {
            >= TableStyle.Dark1 => _table.Style - TableStyle.Dark1 + 1,
            >= TableStyle.Medium1 => _table.Style - TableStyle.Medium1 + 1,
            _ => (int)_table.Style
        };

        return _buffer.TryWrite($"{"TableStyle"u8}{styleCategory}{styleNumber}");
    }

    private readonly bool TryWriteTableStyleInfo()
    {
        return _buffer.TryWrite(
            $"{"\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\""u8}" +
            $"{_table.BandedRows}" +
            $"{"\" showColumnStripes=\""u8}" +
            $"{_table.BandedColumns}" +
            $"{"\"/>"u8}");
    }

    private enum Element
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
}
