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
        ImmutableImage image, // TODO: Should take all images for a sheet here
        CancellationToken token)
    {
        // TODO: Increment number.
        // TODO: Note potential difference between image ID and drawing ID if image is reused across cells.
        var entryName = "xl/drawings/drawing1.xml";
        var entry = archive.CreateEntry(entryName, compressionLevel);
        var writer = new DrawingXml(cellReference, image);
#pragma warning disable EPS06 // Hidden struct copy operation
        return writer.WriteAsync(entry, buffer, token);
#pragma warning restore EPS06 // Hidden struct copy operation
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

    // TODO: cNvPr ID? Should be globally unique. Starts at 0 and increments by 1 for each image (regardless of sheet).
    private static ReadOnlySpan<byte> AnchorEnd => """</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="""u8 + "\""u8;
    private static ReadOnlySpan<byte> ImageIdEnd => "\""u8 + """ name="Image """u8;

    // TODO: Is it OK to use rId when other relation IDs are present? (That are not related to images)
    private static ReadOnlySpan<byte> NameEnd => "\" "u8 + """descr=""></xdr:cNvPr><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId"""u8;

    // TODO: Remove whitespace
    private static ReadOnlySpan<byte> End => "\""u8 + """
                    ></a:blip>
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
    private readonly ImmutableImage _image;
    private Element _next;

    private DrawingXml(SingleCellRelativeReference cellReference, ImmutableImage image)
    {
        _cellReference = cellReference;
        _image = image;
    }

    public bool TryWrite(Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;

        if (_next == Element.Header && !Advance(Header.TryCopyTo(bytes, ref bytesWritten))) return false;
        if (_next == Element.Anchor && !Advance(TryWriteAnchor(bytes, ref bytesWritten))) return false;
        if (_next == Element.Footer && !Advance(TryWriteFooter(bytes, ref bytesWritten))) return false;

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

        // TODO: Support sizing
        // TODO: Subtract offsets
        // TODO: Should they be long?
        // Convert pixels to EMU by multiplying with 9525
        var toColumnOffset = _image.ActualImageWidth * 9525;
        var toRowOffset = _image.ActualImageHeight * 9525;

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

    private readonly bool TryWriteFooter(Span<byte> bytes, ref int bytesWritten)
    {
        var span = bytes.Slice(bytesWritten);
        var written = 0;

        if (!AnchorEnd.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(_image.EmbeddedImageId - 1, span, ref written)) return false;
        if (!ImageIdEnd.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(_image.EmbeddedImageId, span, ref written)) return false;
        if (!NameEnd.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(_image.EmbeddedImageId, span, ref written)) return false; // TODO: Not correct right now: rId should only be unique within the sheet. Starts at 1 and increments for each image in the sheet.
        if (!End.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private enum Element
    {
        Header,
        Anchor,
        Footer,
        Done
    }
}
