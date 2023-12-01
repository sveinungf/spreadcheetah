using SpreadCheetah.Helpers;
using SpreadCheetah.Images;
using SpreadCheetah.Images.Internal;
using System.Diagnostics;

namespace SpreadCheetah.MetadataXml;

internal struct DrawingImageXml
{
    private static ReadOnlySpan<byte> ImageStart => "<xdr:twoCellAnchor editAs=\""u8;
    private static ReadOnlySpan<byte> AnchorStart => "\"><xdr:from><xdr:col>"u8;
    private static ReadOnlySpan<byte> AnchorEnd => """</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="""u8 + "\""u8;
    private static ReadOnlySpan<byte> ImageIdEnd => "\""u8 + """ name="Image """u8;
    private static ReadOnlySpan<byte> NameEnd => "\" "u8 + """descr=""></xdr:cNvPr><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId"""u8;

    // TODO: Bounding box size? (http://officeopenxml.com/drwSp-size.php)
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

        if (_next == Element.ImageStart && !Advance(TryWriteImageStart(bytes, ref bytesWritten))) return false;
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

    private readonly bool TryWriteImageStart(Span<byte> bytes, ref int bytesWritten)
    {
        var span = bytes.Slice(bytesWritten);
        var written = 0;

        if (!ImageStart.TryCopyTo(span, ref written)) return false;

        var moveWithCells = _image.Canvas.Options.HasFlag(ImageCanvasOptions.MoveWithCells);
        var resizeWithCells = _image.Canvas.Options.HasFlag(ImageCanvasOptions.ResizeWithCells);

        var anchorAttribute = (moveWithCells, resizeWithCells) switch
        {
            (true, true) => "twoCell"u8,
            (true, false) => "oneCell"u8,
            (false, false) => "absolute"u8,
            _ => []
        };

        Debug.Assert(!anchorAttribute.IsEmpty);

        if (!anchorAttribute.TryCopyTo(span, ref written)) return false;
        if (!AnchorStart.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private readonly bool TryWriteAnchor(Span<byte> bytes, ref int bytesWritten)
    {
        var span = bytes.Slice(bytesWritten);
        var written = 0;

        var fromColumnOffset = _image.Offset?.Left ?? 0;
        var fromRowOffset = _image.Offset?.Top ?? 0;

        var column = _image.Canvas.FromColumn;
        var row = _image.Canvas.FromRow;

        if (!TryWriteAnchorPart(span, ref written, column, row, fromColumnOffset, fromRowOffset)) return false;
        if (!"</xdr:rowOff></xdr:from><xdr:to><xdr:col>"u8.TryCopyTo(span, ref written)) return false;

        var fillCell = _image.Canvas.Options.HasFlag(ImageCanvasOptions.FillCell);

        var (toColumn, toRow) = fillCell
            ? (_image.Canvas.ToColumn, _image.Canvas.ToRow)
            : (column, row);

        var (actualWidth, actualHeight) = CalculateActualDimensions(_image);
        var (toColumnOffset, toRowOffset) = fillCell
            ? (_image.Offset?.Right ?? 0, _image.Offset?.Bottom ?? 0)
            : (actualWidth + fromColumnOffset, actualHeight + fromRowOffset);

        if (!TryWriteAnchorPart(span, ref written, toColumn, toRow, toColumnOffset, toRowOffset)) return false;

        bytesWritten += written;
        return true;
    }

    private static bool TryWriteAnchorPart(Span<byte> bytes, ref int bytesWritten, ushort column, uint row, int columnPixelsOffset, int rowPixelsOffset)
    {
        if (!SpanHelper.TryWrite(column, bytes, ref bytesWritten)) return false;
        if (!"</xdr:col><xdr:colOff>"u8.TryCopyTo(bytes, ref bytesWritten)) return false;
        if (!SpanHelper.TryWrite(columnPixelsOffset.PixelsToEmu(), bytes, ref bytesWritten)) return false;
        if (!"</xdr:colOff><xdr:row>"u8.TryCopyTo(bytes, ref bytesWritten)) return false;
        if (!SpanHelper.TryWrite(row, bytes, ref bytesWritten)) return false;
        if (!"</xdr:row><xdr:rowOff>"u8.TryCopyTo(bytes, ref bytesWritten)) return false;
        if (!SpanHelper.TryWrite(rowPixelsOffset.PixelsToEmu(), bytes, ref bytesWritten)) return false;
        return true;
    }

    private static (int Width, int Height) CalculateActualDimensions(in WorksheetImage image)
    {
        var canvas = image.Canvas;
        if (canvas.Options.HasFlag(ImageCanvasOptions.Dimensions))
            return (canvas.DimensionWidth, canvas.DimensionHeight);

        var originalDimensions = (image.EmbeddedImage.Width, image.EmbeddedImage.Height);
        return canvas.Options.HasFlag(ImageCanvasOptions.Scaled)
            ? originalDimensions.Scale(canvas.ScaleValue)
            : originalDimensions;
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
