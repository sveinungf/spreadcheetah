using SpreadCheetah.Helpers;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct NumberFormatsXmlPart(
    IList<string>? customNumberFormats,
    SpreadsheetBuffer buffer)
{
    private Element _next;
    private int _nextIndex;
    private string? _currentXmlEncodedFormat;
    private int _currentXmlEncodedFormatIndex;

#pragma warning disable EPS12 // Member can be made readonly (false positive)
    public bool TryWrite()
#pragma warning restore EPS12 // Member can be made readonly (false positive)
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

        return buffer.TryWrite(
            $"{"<numFmts count=\""u8}" +
            $"{formats.Count}" +
            $"{"\">"u8}");
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

            if (_currentXmlEncodedFormat is null)
            {
                const int customFormatStartId = 165;
                var id = _nextIndex + customFormatStartId;

                if (!"<numFmt numFmtId=\""u8.TryCopyTo(span, ref written)) return false;
                if (!SpanHelper.TryWrite(id, span, ref written)) return false;
                if (!"\" formatCode=\""u8.TryCopyTo(span, ref written)) return false;

                _currentXmlEncodedFormat = XmlUtility.XmlEncode(format);
                _currentXmlEncodedFormatIndex = 0;
            }

            if (!SpanHelper.TryWriteLongString(_currentXmlEncodedFormat, ref _currentXmlEncodedFormatIndex, span, ref written))
            {
                buffer.Advance(written);
                return false;
            }

            if (!"\"/>"u8.TryCopyTo(span, ref written)) return false;

            buffer.Advance(written);

            _currentXmlEncodedFormat = null;
            _currentXmlEncodedFormatIndex = 0;
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
