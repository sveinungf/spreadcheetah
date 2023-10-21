using SpreadCheetah.Helpers;
using SpreadCheetah.Images.Internal;

namespace SpreadCheetah.MetadataXml;

internal struct DrawingImageXml
{
    // TODO: twoCellAnchor?
    // TODO: Parameter for editAs
    private static ReadOnlySpan<byte> ImageStart => """<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>"""u8 + "\""u8;

    // TODO: cNvPr ID? Should be globally unique. Starts at 0 and increments by 1 for each image (regardless of sheet).
    private static ReadOnlySpan<byte> AnchorEnd => """</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="""u8 + "\""u8;
    private static ReadOnlySpan<byte> ImageIdEnd => "\""u8 + """ name="Image """u8;

    // TODO: Is it OK to use rId when other relation IDs are present? (That are not related to images)
    private static ReadOnlySpan<byte> NameEnd => "\" "u8 + """descr=""></xdr:cNvPr><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId"""u8;

    private static ReadOnlySpan<byte> ImageEnd => "\""u8 +
        """></a:blip><a:stretch/></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="1" y="1"/><a:ext cx="1" cy="1"/>"""u8 +
        """</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="0"><a:noFill/></a:ln></xdr:spPr>"""u8 +
        """</xdr:pic><xdr:clientData/></xdr:twoCellAnchor>"""u8;

    private readonly WorksheetImage _image;
    private Element _next;

    public DrawingImageXml(WorksheetImage image)
    {
        _image = image;
    }

    public bool TryWrite(Span<byte> bytes, out int bytesWritten)
    {
        bytesWritten = 0;

        if (_next == Element.ImageStart && !Advance(ImageStart.TryCopyTo(bytes, ref bytesWritten))) return false;
        if (_next == Element.Anchor && !Advance(TryWriteAnchor(bytes, ref bytesWritten))) return false;
        if (_next == Element.ImageEnd && !Advance(TryWriteImageEnd(bytes, ref bytesWritten))) return false;

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
        var toColumnOffset = _image.Image.ActualImageWidth * 9525;
        var toRowOffset = _image.Image.ActualImageHeight * 9525;

        var column = _image.Reference.Column - 1;
        var row = _image.Reference.Row - 1;

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

    private readonly bool TryWriteImageEnd(Span<byte> bytes, ref int bytesWritten)
    {
        var span = bytes.Slice(bytesWritten);
        var written = 0;

        if (!AnchorEnd.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(_image.Image.EmbeddedImageId - 1, span, ref written)) return false;
        if (!ImageIdEnd.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(_image.Image.EmbeddedImageId, span, ref written)) return false;
        if (!NameEnd.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(_image.Image.EmbeddedImageId, span, ref written)) return false; // TODO: Not correct right now: rId should only be unique within the sheet. Starts at 1 and increments for each image in the sheet.
        if (!ImageEnd.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private enum Element
    {
        ImageStart,
        Anchor,
        ImageEnd,
        Done
    }
}
