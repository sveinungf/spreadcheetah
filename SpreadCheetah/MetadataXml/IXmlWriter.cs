namespace SpreadCheetah.MetadataXml;

internal interface IXmlWriter
{
    bool TryWrite(Span<byte> bytes, out int bytesWritten);
}
