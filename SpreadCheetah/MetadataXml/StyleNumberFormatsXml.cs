using SpreadCheetah.Helpers;

namespace SpreadCheetah.MetadataXml;

internal struct StyleNumberFormatsXml
{
    private readonly List<KeyValuePair<string, int>>? _customNumberFormats;
    private Element _next;
    private int _nextIndex;

    public StyleNumberFormatsXml(List<KeyValuePair<string, int>>? customNumberFormats)
    {
        _customNumberFormats = customNumberFormats;
    }

    public bool TryWrite(Span<byte> bytes, ref int bytesWritten)
    {
        if (_next == Element.Header && !Advance(TryWriteHeader(bytes, ref bytesWritten))) return false;
        if (_next == Element.NumberFormats && !Advance(TryWriteNumberFormats(bytes, ref bytesWritten))) return false;
        if (_next == Element.Footer && !Advance(TryWriteFooter(bytes, ref bytesWritten))) return false;

        return true;
    }

    private bool Advance(bool success)
    {
        if (success)
            ++_next;

        return success;
    }

    private readonly bool TryWriteHeader(Span<byte> bytes, ref int bytesWritten)
    {
        if (_customNumberFormats is not { } formats)
            return """<numFmts count="0"/>"""u8.TryCopyTo(bytes, ref bytesWritten);

        var span = bytes.Slice(bytesWritten);
        var written = 0;

        if (!"<numFmts count=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(formats.Count, span, ref written)) return false;
        if (!"\">"u8.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private bool TryWriteNumberFormats(Span<byte> bytes, ref int bytesWritten)
    {
        if (_customNumberFormats is not { } formats)
            return true;

        for (; _nextIndex < formats.Count; ++_nextIndex)
        {
            var format = formats[_nextIndex];
            var span = bytes.Slice(bytesWritten);
            var written = 0;

            if (!"<numFmt numFmtId=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(format.Value, span, ref written)) return false;
            if (!"\" formatCode=\""u8.TryCopyTo(span, ref written)) return false;

            var numberFormat = XmlUtility.XmlEncode(format.Key);
            if (!SpanHelper.TryWrite(numberFormat, span, ref written)) return false;
            if (!"\"/>"u8.TryCopyTo(span, ref written)) return false;

            bytesWritten += written;
        }

        return true;
    }

    private readonly bool TryWriteFooter(Span<byte> bytes, ref int bytesWritten)
        => _customNumberFormats is null || "</numFmts>"u8.TryCopyTo(bytes, ref bytesWritten);

    private enum Element
    {
        Header,
        NumberFormats,
        Footer,
        Done
    }
}
