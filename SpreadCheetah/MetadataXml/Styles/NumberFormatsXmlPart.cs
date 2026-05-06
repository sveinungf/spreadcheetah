namespace SpreadCheetah.MetadataXml.Styles;

internal struct NumberFormatsXmlPart(
    IList<string>? customNumberFormats,
    SpreadsheetBuffer buffer)
{
    private Element _next;
    private int _nextIndex;
    private NumberFormatXmlPart? _currentFormatPart;

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
            if (_currentFormatPart is not { } current)
            {
                const int customFormatStartId = 165;
                var id = _nextIndex + customFormatStartId;
                var format = formats[_nextIndex];
                current = new NumberFormatXmlPart(id, format);
            }

            if (!current.TryWrite(buffer))
            {
                _currentFormatPart = current;
                return false;
            }

            _currentFormatPart = null;
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
