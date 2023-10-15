using SpreadCheetah.CellReferences;
using SpreadCheetah.Helpers;
using SpreadCheetah.Images.Internal;
using System.IO.Compression;
using System.Net;

namespace SpreadCheetah.MetadataXml;

internal struct DrawingXml : IXmlWriter
{
    public static ValueTask WriteAsync(
        ZipArchive archive,
        CompressionLevel compressionLevel,
        SpreadsheetBuffer buffer,
        SingleCellRelativeReference cellReference,
        ImmutableImageOptions imageOptions,
        CancellationToken token)
    {
        // TODO: Increment number
        var entryName = "xl/drawings/drawing1.xml";
        var entry = archive.CreateEntry(entryName, compressionLevel);
        var writer = new DrawingXml(cellReference, imageOptions);
        return writer.WriteAsync(entry, buffer, token);
    }

    // TODO: twoCellAnchor?
    // TODO: Parameter for editAs
    // TODO: Remove whitespace
    private static ReadOnlySpan<byte> Header => """
        <?xml version="1.0" encoding="utf-8"?>
        <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <xdr:twoCellAnchor editAs="oneCell">
                <xdr:from>
                    <xdr:col>
        """u8 + "\""u8;

    // TODO: cNvPr ID?
    private static ReadOnlySpan<byte> AnchorEnd => """</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="0" name="""u8;

    // TODO: rId1 - Should be correct ID
    // TODO: Remove whitespace
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

    private readonly SingleCellRelativeReference _cellReference;
    private readonly ImmutableImageOptions _imageOptions;
    private readonly string _xmlEncodedName;
    private int _xmlEncodedNameIndex;
    private Element _next;

    private DrawingXml(SingleCellRelativeReference cellReference, ImmutableImageOptions imageOptions)
    {
        _cellReference = cellReference;
        _imageOptions = imageOptions;
        _xmlEncodedName = WebUtility.HtmlEncode(imageOptions.Name);
    }

    public bool TryWrite(Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;

        if (_next == Element.Header && !Advance(Header.TryCopyTo(bytes, ref bytesWritten))) return false;
        if (_next == Element.Anchor && !Advance(TryWriteAnchor(bytes, ref bytesWritten))) return false;
        if (_next == Element.AnchorEnd && !Advance(AnchorEnd.TryCopyTo(bytes, ref bytesWritten))) return false;
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

    private readonly bool TryWriteAnchor(Span<byte> bytes, ref int bytesWritten)
    {
        var span = bytes.Slice(bytesWritten);
        var written = 0;

        // TODO: Upper-left offsets
        // TODO: Should they be long?
        const int fromColumnOffset = 0;
        const int fromRowOffset = 0;

        // TODO: Should be calculated from image size minus any offset
        // TODO: Should they be long?
        const int toColumnOffset = 406080;
        const int toRowOffset = 81360;

        var column = _cellReference.Column - 1;
        var row = _cellReference.Row - 1;

        if (!TryWriteAnchorPart(span, ref written, column, row, fromColumnOffset, fromRowOffset)) return false;
        if (!"</xdr:rowOff></xdr:from><xdr:to><xdr:col>"u8.TryCopyTo(bytes, ref written)) return false;
        if (!TryWriteAnchorPart(span, ref written, column, row, toColumnOffset, toRowOffset)) return false;

        bytesWritten += written;
        return true;

        static bool TryWriteAnchorPart(Span<byte> bytes, ref int bytesWritten, int column, int row, int columnOffset, int rowOffset)
        {
            if (!SpanHelper.TryWrite(column, bytes, ref bytesWritten)) return false;
            if (!"</xdr:col><xdr:colOff"u8.TryCopyTo(bytes, ref bytesWritten)) return false;
            if (!SpanHelper.TryWrite(columnOffset, bytes, ref bytesWritten)) return false;
            if (!"</xdr:colOff><xdr:row>"u8.TryCopyTo(bytes, ref bytesWritten)) return false;
            if (!SpanHelper.TryWrite(row, bytes, ref bytesWritten)) return false;
            if (!"</xdr:row><xdr:rowOff>"u8.TryCopyTo(bytes, ref bytesWritten)) return false;
            if (!SpanHelper.TryWrite(rowOffset, bytes, ref bytesWritten)) return false;
            return true;
        }
    }

    private bool TryWriteName(Span<byte> bytes, ref int bytesWritten)
    {
        return SpanHelper.TryWriteLongString(_xmlEncodedName, ref _xmlEncodedNameIndex, bytes, ref bytesWritten);
    }

    private enum Element
    {
        Header,
        Anchor,
        AnchorEnd,
        Name,
        Footer,
        Done
    }
}
