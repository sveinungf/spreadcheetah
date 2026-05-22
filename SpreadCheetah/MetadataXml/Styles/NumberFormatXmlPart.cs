namespace SpreadCheetah.MetadataXml.Styles;

internal struct NumberFormatXmlPart(
    int id,
    string format)
{
    private int? _currentIndex;

    public bool TryWrite(SpreadsheetBuffer buffer)
    {
        if (_currentIndex is not { } index)
        {
            if (!buffer.TryWrite($"{"<numFmt numFmtId=\""u8}{id}{"\" formatCode=\""u8}"))
                return false;

            index = 0;
        }

        if (buffer.WriteLongString(format, ref index) && buffer.TryWrite("\"/>"u8))
            return true;

        _currentIndex = index;
        return false;
    }
}