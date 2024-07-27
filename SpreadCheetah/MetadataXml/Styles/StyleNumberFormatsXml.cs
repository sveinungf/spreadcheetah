using SpreadCheetah.Helpers;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct StyleNumberFormatsXml(
    List<KeyValuePair<string, int>>? customNumberFormats,
    SpreadsheetBuffer buffer)
{
    private Element _next;
    private int _nextIndex;

    public bool TryWrite()
    {
        while (MoveNext())
        {
            if (!Current)
                return false;
        }

        return true;
    }

    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.Header => TryWriteHeader(),
            Element.NumberFormats => TryWriteNumberFormats(),
            _ => TryWriteFooter()
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteHeader()
    {
        if (customNumberFormats is not { } formats)
            return buffer.TryWrite("""<numFmts count="0"/>"""u8);

        var span = buffer.GetSpan();
        var written = 0;

        if (!"<numFmts count=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(formats.Count, span, ref written)) return false;
        if (!"\">"u8.TryCopyTo(span, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    private bool TryWriteNumberFormats()
    {
        if (customNumberFormats is not { } formats)
            return true;

        for (; _nextIndex < formats.Count; ++_nextIndex)
        {
            var format = formats[_nextIndex];
            var span = buffer.GetSpan();
            var written = 0;

            if (!"<numFmt numFmtId=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(format.Value, span, ref written)) return false;
            if (!"\" formatCode=\""u8.TryCopyTo(span, ref written)) return false;

            // TODO: This can be longer than the minimum buffer length. Should use the "long string" approach
            var numberFormat = XmlUtility.XmlEncode(format.Key);
            if (!SpanHelper.TryWrite(numberFormat, span, ref written)) return false;
            if (!"\"/>"u8.TryCopyTo(span, ref written)) return false;

            buffer.Advance(written);
        }

        return true;
    }

    private readonly bool TryWriteFooter()
        => customNumberFormats is null || buffer.TryWrite("</numFmts>"u8);

    private enum Element
    {
        Header,
        NumberFormats,
        Footer,
        Done
    }
}
