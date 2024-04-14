namespace SpreadCheetah.MetadataXml;

internal interface IBufferXmlWriter
{
    bool TryWrite(SpreadsheetBuffer buffer);
}
