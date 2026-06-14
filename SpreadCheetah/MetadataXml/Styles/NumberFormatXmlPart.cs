using SpreadCheetah.Helpers;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct NumberFormatXmlPart(
    int id,
    string format)
{
    private BufferWriteProgress _progress;

    public bool TryWrite(SpreadsheetBuffer buffer)
    {
        return buffer.TryWrite2(
            _progress, out _progress,
            $"{"<numFmt numFmtId=\""u8}{id}{"\" formatCode=\""u8}{format}{"\"/>"u8}");
    }
}