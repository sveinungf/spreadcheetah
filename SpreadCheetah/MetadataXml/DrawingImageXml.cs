using SpreadCheetah.Helpers;
using SpreadCheetah.Images.Internal;

namespace SpreadCheetah.MetadataXml;

internal struct DrawingImageXml
{
    // TODO: twoCellAnchor?
    // TODO: Parameter for editAs (oneCell/twoCell/absolute)
    private static ReadOnlySpan<byte> ImageStart => """<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>"""u8;
    private static ReadOnlySpan<byte> AnchorEnd => """</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="""u8 + "\""u8;
    private static ReadOnlySpan<byte> ImageIdEnd => "\""u8 + """ name="Image """u8;
    private static ReadOnlySpan<byte> NameEnd => "\" "u8 + """descr=""></xdr:cNvPr><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId"""u8;
    private static ReadOnlySpan<byte> ImageEnd => "\""u8 +
        """></a:blip><a:stretch/></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="1" y="1"/><a:ext cx="1" cy="1"/>"""u8 +
        """</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="0"><a:noFill/></a:ln></xdr:spPr>"""u8 +
        """</xdr:pic><xdr:clientData/></xdr:twoCellAnchor>"""u8;

    private readonly WorksheetImage _image;
    private readonly int _worksheetImageIndex;
    private Element _next;

    public DrawingImageXml(WorksheetImage image, int worksheetImageIndex)
    {
        _image = image;
        _worksheetImageIndex = worksheetImageIndex;
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
        const int fromColumnOffset = 0;
        const int fromRowOffset = 0;

        var column = _image.Reference.Column - 1;
        var row = _image.Reference.Row - 1;

        if (!TryWriteAnchorPart(span, ref written, column, row, fromColumnOffset, fromRowOffset)) return false;
        if (!"</xdr:rowOff></xdr:from><xdr:to><xdr:col>"u8.TryCopyTo(span, ref written)) return false;

        // TODO: Support sizing
        // TODO: Subtract offsets
        // Convert pixels to EMU by multiplying with 9525
        var toColumnOffset = _image.Image.ActualImageWidth * 9525;
        var toRowOffset = _image.Image.ActualImageHeight * 9525;

        if (!TryWriteAnchorPart(span, ref written, column, row, toColumnOffset, toRowOffset)) return false;

        bytesWritten += written;
        return true;

        static bool TryWriteAnchorPart(Span<byte> bytes, ref int bytesWritten, int column, int row, int columnOffset, int rowOffset)
        {
            if (!SpanHelper.TryWrite(column, bytes, ref bytesWritten)) return false;
            if (!"</xdr:col><xdr:colOff>"u8.TryCopyTo(bytes, ref bytesWritten)) return false;
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
        if (!SpanHelper.TryWrite(_worksheetImageIndex, span, ref written)) return false;
        if (!ImageIdEnd.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(_image.ImageNumber, span, ref written)) return false;
        if (!NameEnd.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(_worksheetImageIndex + 1, span, ref written)) return false;
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
