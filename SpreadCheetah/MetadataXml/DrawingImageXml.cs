using SpreadCheetah.Helpers;
using SpreadCheetah.Images;
using SpreadCheetah.Images.Internal;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace SpreadCheetah.MetadataXml;

[StructLayout(LayoutKind.Auto)]
internal struct DrawingImageXml(
    WorksheetImage image,
    int worksheetImageIndex,
    SpreadsheetBuffer buffer)
{
    private static ReadOnlySpan<byte> ImageStart => "<xdr:twoCellAnchor editAs=\""u8;
    private static ReadOnlySpan<byte> AnchorStart => "\"><xdr:from><xdr:col>"u8;
    private static ReadOnlySpan<byte> AnchorEnd => """</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="""u8 + "\""u8;
    private static ReadOnlySpan<byte> ImageIdEnd => "\""u8 + """ name="Image """u8;
    private static ReadOnlySpan<byte> NameEnd => "\" "u8 + """descr=""></xdr:cNvPr><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId"""u8;

    private static ReadOnlySpan<byte> ImageEnd => "\""u8 +
        """></a:blip><a:stretch/></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="1" y="1"/><a:ext cx="1" cy="1"/>"""u8 +
        """</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="0"><a:noFill/></a:ln></xdr:spPr>"""u8 +
        """</xdr:pic><xdr:clientData/></xdr:twoCellAnchor>"""u8;

    private Element _next;

#pragma warning disable EPS12 // A struct member can be made readonly
    public bool TryWrite()
#pragma warning restore EPS12 // A struct member can be made readonly
    {
        while (MoveNext())
        {
            if (!Current)
                return false;
        }

        return true;
    }

    public bool Current { get; private set; }

    public bool MoveNext()
    {
        Current = _next switch
        {
            Element.ImageStart => TryWriteImageStart(),
            Element.Anchor => TryWriteAnchor(),
            _ => TryWriteImageEnd()
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteImageStart()
    {
        var span = buffer.GetSpan();
        var written = 0;

        if (!ImageStart.TryCopyTo(span, ref written)) return false;

        var moveWithCells = image.Canvas.Options.HasFlag(ImageCanvasOptions.MoveWithCells);
        var resizeWithCells = image.Canvas.Options.HasFlag(ImageCanvasOptions.ResizeWithCells);

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

        buffer.Advance(written);
        return true;
    }

    private readonly bool TryWriteAnchor()
    {
        var span = buffer.GetSpan();
        var written = 0;

        var fromColumnOffset = image.Offset?.Left ?? 0;
        var fromRowOffset = image.Offset?.Top ?? 0;

        var column = image.Canvas.FromColumn;
        var row = image.Canvas.FromRow;

        if (!TryWriteAnchorPart(span, ref written, column, row, fromColumnOffset, fromRowOffset)) return false;
        if (!"</xdr:rowOff></xdr:from><xdr:to><xdr:col>"u8.TryCopyTo(span, ref written)) return false;

        var fillCell = image.Canvas.Options.HasFlag(ImageCanvasOptions.FillCell);

        var (toColumn, toRow) = fillCell
            ? (image.Canvas.ToColumn, image.Canvas.ToRow)
            : (column, row);

        var (actualWidth, actualHeight) = CalculateActualDimensions(image);
        var (toColumnOffset, toRowOffset) = fillCell
            ? (image.Offset?.Right ?? 0, image.Offset?.Bottom ?? 0)
            : (actualWidth + fromColumnOffset, actualHeight + fromRowOffset);

        if (!TryWriteAnchorPart(span, ref written, toColumn, toRow, toColumnOffset, toRowOffset)) return false;

        buffer.Advance(written);
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

    private readonly bool TryWriteImageEnd()
    {
        var span = buffer.GetSpan();
        var written = 0;

        if (!AnchorEnd.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(worksheetImageIndex, span, ref written)) return false;
        if (!ImageIdEnd.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(image.ImageNumber, span, ref written)) return false;
        if (!NameEnd.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(worksheetImageIndex + 1, span, ref written)) return false;
        if (!ImageEnd.TryCopyTo(span, ref written)) return false;

        buffer.Advance(written);
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
