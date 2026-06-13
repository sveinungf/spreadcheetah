namespace SpreadCheetah.MetadataXml.Styles;

internal struct NumberFormatXmlPart(
    int id,
    string format)
{
    private int _step;
    private int _index;

    public bool TryWrite(SpreadsheetBuffer buffer)
    {
        //if (_currentIndex is not { } index)
        //{
        //    if (!buffer.TryWrite2(0, 0, $"{"<numFmt numFmtId=\""u8}{id}{"\" formatCode=\""u8}", out var step, out var index2))
        //        return false;

        //    index = 0;
        //}

        //if (buffer.WriteLongString(format, ref index) && buffer.TryWrite("\"/>"u8))
        //    return true;

        //_currentIndex = index;
        //return false;

        return buffer.TryWrite2(
            _step, _index,
            $"{"<numFmt numFmtId=\""u8}{id}{"\" formatCode=\""u8}{format}{"\"/>"u8}",
            out _step, out _index);
    }
}