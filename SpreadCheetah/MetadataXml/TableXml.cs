namespace SpreadCheetah.MetadataXml;

internal struct TableXml
{
    private static ReadOnlySpan<byte> Header =>
        """<?xml version="1.0" encoding="utf-8"?>"""u8 +
        """<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" """u8 +
        """id="1" name="Table1" displayName="Table1" ref="A1:A2" insertRow="1" totalsRowShown="0">"""u8 +
        """<autoFilter ref="A1:A2"/>"""u8 +
        """<tableColumns count="1">"""u8 +
        """<tableColumn id="1" name="Column1"/>"""u8 +
        """</tableColumns>"""u8 +
        """<tableStyleInfo name="TableStyleLight1" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>"""u8;

    private readonly SpreadsheetBuffer _buffer;
    private Element _next;

    private TableXml(SpreadsheetBuffer buffer)
    {
        _buffer = buffer;
    }

    public readonly TableXml GetEnumerator() => this;
    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => _buffer.TryWrite(Header),
            _ => _buffer.TryWrite("</table>"u8)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private enum Element
    {
        Header,
        Footer,
        Done
    }
}
