using SpreadCheetah.Helpers;
using System.IO.Compression;
using System.Net;

namespace SpreadCheetah.MetadataXml;

internal struct DrawingXml : IXmlWriter
{
    public static ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        string name,
        CancellationToken token)
    {
        // TODO: Increment number
        var entryName = "xl/drawings/drawing1.xml";
        var entry = archive.CreateEntry(entryName, compressionLevel);
        var writer = new DrawingXml(name);
        return writer.WriteAsync(entry, buffer, token);
    }

    // TODO: twoCellAnchor?
    // TODO: Parameter for editAs
    // TODO: col/row
    // TODO: Offsets
    // TODO: Remove whitespace
    private static ReadOnlySpan<byte> Header => """
        <?xml version="1.0" encoding="utf-8"?>
        <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <xdr:twoCellAnchor editAs="oneCell">
                <xdr:from>
                    <xdr:col>1</xdr:col>
                    <xdr:colOff>0</xdr:colOff>
                    <xdr:row>1</xdr:row>
                    <xdr:rowOff>0</xdr:rowOff>
                </xdr:from>
                <xdr:to>
                    <xdr:col>2</xdr:col>
                    <xdr:colOff>406080</xdr:colOff>
                    <xdr:row>8</xdr:row>
                    <xdr:rowOff>81360</xdr:rowOff>
                </xdr:to>
                <xdr:pic>
                    <xdr:nvPicPr>
                        <xdr:cNvPr id="0" name=
        """u8 + "\""u8;

    // TODO: rId1?
    private static ReadOnlySpan<byte> Footer => "\""u8 + """
                    descr=""></xdr:cNvPr>
                        <xdr:cNvPicPr/>
                    </xdr:nvPicPr>
                    <xdr:blipFill>
                        <a:blip r:embed="rId1"></a:blip>
                        <a:stretch/>
                    </xdr:blipFill>
                    <xdr:spPr>
                        <a:xfrm>
                            <a:off x="1" y="1"/>
                            <a:ext cx="1" cy="1"/>
                        </a:xfrm>
                        <a:prstGeom prst="rect">
                            <a:avLst/>
                        </a:prstGeom>
                        <a:ln w="0">
                            <a:noFill/>
                        </a:ln>
                    </xdr:spPr>
                </xdr:pic>
                <xdr:clientData/>
            </xdr:twoCellAnchor>
        </xdr:wsDr>
        """u8;

    private readonly string _xmlEncodedName;
    private int _xmlEncodedNameIndex;
    private Element _next;

    private DrawingXml(string name)
    {
        _xmlEncodedName = WebUtility.HtmlEncode(name);
    }

    public bool TryWrite(Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;

        if (_next == Element.Header && !Advance(Header.TryCopyTo(bytes, ref bytesWritten))) return false;
        if (_next == Element.Name && !Advance(TryWriteName(bytes, ref bytesWritten))) return false;
        if (_next == Element.Footer && !Advance(Footer.TryCopyTo(bytes, ref bytesWritten))) return false;

        return true;
    }

    private bool Advance(bool success)
    {
        if (success)
            ++_next;

        return success;
    }

    private bool TryWriteName(Span<byte> bytes, ref int bytesWritten)
    {
        return SpanHelper.TryWriteLongString(_xmlEncodedName, ref _xmlEncodedNameIndex, bytes, ref bytesWritten);
    }

    private enum Element
    {
        Header,
        Name,
        Footer,
        Done
    }
}
