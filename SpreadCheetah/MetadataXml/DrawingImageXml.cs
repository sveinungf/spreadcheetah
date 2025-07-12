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
    Worksheet worksheet,
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
            Element.FromAnchor => TryWriteFromAnchor(),
            Element.BetweenAnchors => buffer.TryWrite("</xdr:rowOff></xdr:from><xdr:to><xdr:col>"u8),
            Element.ToAnchor => TryWriteToAnchor(),
            _ => TryWriteImageEnd()
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteImageStart()
    {
        var moveWithCells = image.Canvas.Options.HasFlag(ImageCanvasOptions.MoveWithCells);
        var resizeWithCells = image.Canvas.Options.HasFlag(ImageCanvasOptions.ResizeWithCells);

        Debug.Assert(moveWithCells || !resizeWithCells);
        var anchorAttribute = (moveWithCells, resizeWithCells) switch
        {
            (true, true) => "twoCell"u8,
            (true, false) => "oneCell"u8,
            _ => "absolute"u8
        };

        return buffer.TryWrite($"{ImageStart}{anchorAttribute}{AnchorStart}");
    }

    private readonly bool TryWriteFromAnchor()
    {
        var fromColumnOffset = image.Offset?.Left ?? 0;
        var fromRowOffset = image.Offset?.Top ?? 0;

        var column = image.Canvas.FromColumn;
        var row = image.Canvas.FromRow;

        return TryWriteAnchorPart(column, row, fromColumnOffset.PixelsToOffset(), fromRowOffset.PixelsToOffset());
    }

    private readonly bool TryWriteToAnchor()
    {
        var column = image.Canvas.FromColumn;
        var row = image.Canvas.FromRow;

        var (actualWidth, actualHeight) = CalculateActualDimensions(image);
        var fillCell = image.Canvas.Options.HasFlag(ImageCanvasOptions.FillCell);

        var (columnCount, toColumnOffset) = fillCell
            ? (image.Canvas.ColumnCount, image.Offset?.Right ?? 0)
            : CalculateColumns(column, actualWidth);

        var (rowCount, toRowOffset) = fillCell
            ? (image.Canvas.RowCount, image.Offset?.Bottom ?? 0)
            : CalculateRows((int)row, actualHeight);

        var toColumn = (ushort)(column + columnCount);
        var toRow = row + rowCount;

        return TryWriteAnchorPart(toColumn, toRow, toColumnOffset.PixelsToOffset(), toRowOffset.PixelsToOffset());
    }

    private readonly (ushort ColumnCount, int ToColumnOffsetInPixels) CalculateColumns(int fromColumn, int imageWidth)
    {
        var fromColumnOffsetInPixels = image.Offset?.Left ?? 0;
        double remainingWidthInPixels = fromColumnOffsetInPixels + imageWidth;
        var toColumnOffsetInPixels = remainingWidthInPixels;
        var columnWidths = worksheet.ColumnWidthRuns.GetColumnWidths(worksheet.DefaultColumnWidth);
        ushort columnCount = 0;

        foreach (var columnWidth in columnWidths.Skip(fromColumn))
        {
            Debug.Assert(columnWidth >= 0, "Column width must be non-negative.");

            var widthInPixels = columnWidth * 9;
            remainingWidthInPixels -= widthInPixels;
            if (remainingWidthInPixels < 0)
                break;

            columnCount++;
            toColumnOffsetInPixels = remainingWidthInPixels;
        }

        return (columnCount, Round(toColumnOffsetInPixels));
    }

    private readonly (uint RowCount, int ToRowOffsetInPixels) CalculateRows(int fromRow, int imageHeight)
    {
        var fromRowOffsetInPixels = image.Offset?.Top ?? 0;
        var remainingHeightInPixels = fromRowOffsetInPixels + imageHeight;
        var toRowOffsetInPixels = remainingHeightInPixels;
        var rowHeights = worksheet.RowHeightRuns.GetRowHeights();
        var rowCount = 0u;

        foreach (var rowHeight in rowHeights.Skip(fromRow))
        {
            Debug.Assert(rowHeight >= 0, "Row height must be non-negative.");

            const double heightToPixels = 4.0 / 3.0;
            var heightInPixels = (rowHeight * heightToPixels) - 0.5;
            remainingHeightInPixels -= Round(heightInPixels);
            if (remainingHeightInPixels < 0)
                break;

            rowCount++;
            toRowOffsetInPixels = remainingHeightInPixels;
        }

        return (rowCount, toRowOffsetInPixels);
    }

    private static int Round(double value) => (int)Math.Round(value, MidpointRounding.AwayFromZero);

    private readonly bool TryWriteAnchorPart(ushort column, uint row, int columnOffset, int rowOffset)
    {
        return buffer.TryWrite(
            $"{column}" +
            $"{"</xdr:col><xdr:colOff>"u8}" +
            $"{columnOffset}" +
            $"{"</xdr:colOff><xdr:row>"u8}" +
            $"{row}" +
            $"{"</xdr:row><xdr:rowOff>"u8}" +
            $"{rowOffset}");
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
        return buffer.TryWrite(
            $"{AnchorEnd}" +
            $"{worksheetImageIndex}" +
            $"{ImageIdEnd}" +
            $"{image.ImageNumber}" +
            $"{NameEnd}" +
            $"{worksheetImageIndex + 1}" +
            $"{ImageEnd}");
    }

    private enum Element
    {
        ImageStart,
        FromAnchor,
        BetweenAnchors,
        ToAnchor,
        ImageEnd,
        Done
    }
}
