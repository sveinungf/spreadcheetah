using SpreadCheetah.CellReferences;
using SpreadCheetah.Tables;

namespace SpreadCheetah.MetadataXml;

internal struct TableXml
{
    // TODO: "id" should be globally unique?
    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" """u8 +
        """id="1" name="""u8 + "\""u8;

    // TODO: Set value for "insertRow"?
    // TODO: Set value for "totalsRowShown"
    // TODO: Consider having autoFilter by default for all columns
    private static ReadOnlySpan<byte> ReferenceEnd => "\""u8 +
        """ insertRow="1" totalsRowShown="0">"""u8 +
        """<autoFilter ref="A1:A2"/>"""u8 +
        """<tableColumns count="1">"""u8 +
        """<tableColumn id="1" name="Column1"/>"""u8 +
        """</tableColumns>"""u8 +
        """<tableStyleInfo name="TableStyleLight1" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>"""u8;

    private readonly SpreadsheetBuffer _buffer;
    private readonly Table _table;
    private readonly ReadOnlyMemory<string?> _headerRowNames; // TODO: Set value. Might be empty.
    private readonly ushort _startColumn; // TODO: Set value
    private readonly uint _startRow; // TODO: Set value
    private readonly uint _endRow; // TODO: Set value
    private Element _next;

    private TableXml(Table table, SpreadsheetBuffer buffer)
    {
        _buffer = buffer;
        _table = table;
    }

    public readonly TableXml GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => _buffer.TryWrite(Header),
            Element.Name => _buffer.TryWrite($"{_table.Name}"), // TODO: Can it be longer than minimum buffer length? Maybe it should be disallowed?
            Element.NameEnd => _buffer.TryWrite("\" displayName=\""u8),
            Element.DisplayName => _buffer.TryWrite($"{_table.Name}"), // TODO: Can it be longer than minimum buffer length? Maybe it should be disallowed?
            Element.DisplayNameEnd => _buffer.TryWrite("\" ref=\""u8),
            Element.Reference => TryWriteCellRangeReference(),
            Element.ReferenceEnd => _buffer.TryWrite(ReferenceEnd),
            _ => _buffer.TryWrite("</table>"u8)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteCellRangeReference()
    {
        var fromCell = new SimpleSingleCellReference(_startColumn, _startRow);

        // TODO: Can columnCount end up being "0"?
        var columnCount = _table.NumberOfColumns ?? _headerRowNames.Length;
        var endColumn = (ushort)(_startColumn + columnCount);
        var toCell = new SimpleSingleCellReference(endColumn, _endRow);

        return _buffer.TryWrite($"{fromCell}:{toCell}");
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
        Footer,
        Done
    }
}
