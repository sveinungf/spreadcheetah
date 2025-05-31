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

    private readonly bool TryWriteAnchor()
    {
        var span = buffer.GetSpan();
        var written = 0;

        var fromColumnOffset = image.Offset?.Left.PixelsToEmu() ?? 0;
        var fromRowOffset = image.Offset?.Top.PixelsToEmu() ?? 0;

        var column = image.Canvas.FromColumn;
        var row = image.Canvas.FromRow;

        if (!TryWriteAnchorPart(span, ref written, column, row, fromColumnOffset, fromRowOffset)) return false;
        if (!"</xdr:rowOff></xdr:from><xdr:to><xdr:col>"u8.TryCopyTo(span, ref written)) return false;

        var (actualWidth, actualHeight) = CalculateActualDimensions(image);
        var fillCell = image.Canvas.Options.HasFlag(ImageCanvasOptions.FillCell);

        var (columnCount, toColumnOffset) = fillCell
            ? (0, image.Offset?.Right.PixelsToEmu() ?? 0)
            : CalculateColumns(fromColumnOffset, actualWidth, GetColumnWidths());

        var (rowCount, toRowOffset) = fillCell
            ? (0, image.Offset?.Bottom.PixelsToEmu() ?? 0)
            : CalculateRows(fromRowOffset, actualHeight, GetRowHeights());

        var toColumn = (ushort)(column + columnCount);
        var toRow = (uint)(row + rowCount);

        if (!TryWriteAnchorPart(span, ref written, toColumn, toRow, toColumnOffset, toRowOffset)) return false;

        buffer.Advance(written);
        return true;
    }

    private static IEnumerable<double> GetColumnWidths()
    {
        for (var i = 0; i < SpreadsheetConstants.MaxNumberOfColumns; ++i)
        {
            yield return SpreadsheetConstants.DefaultColumnWidthInEmu;
        }
    }

    private static IEnumerable<double> GetRowHeights()
    {
        for (var i = 0; i < SpreadsheetConstants.MaxNumberOfRows; ++i)
        {
            yield return SpreadsheetConstants.DefaultRowHeightInEmu;
        }
    }

    private static (int ColumnCount, int ToColumnOffset) CalculateColumns(int fromColumnOffset, int imageWidth, IEnumerable<double> columnWidths)
    {
        double remainingWidthInEmu = fromColumnOffset + imageWidth.PixelsToEmu();
        var columnCount = 0;
        var toColumnOffset = remainingWidthInEmu;
        const double delta = 0.776275;
        const double emuPerColumnWidthUnit = 609600 / (SpreadsheetConstants.DefaultColumnWidthInEmu + delta);

        foreach (var columnWidth in columnWidths)
        {
            Debug.Assert(columnWidth >= 0, "Column width must be non-negative.");

            remainingWidthInEmu -= (columnWidth + delta) * emuPerColumnWidthUnit;
            if (remainingWidthInEmu < 0)
                return (columnCount, (int)toColumnOffset);

            columnCount++;
            toColumnOffset = remainingWidthInEmu;
        }

        // TODO: Unreachable?
        return (columnCount, (int)toColumnOffset);
    }

    private static (int RowCount, int ToRowOffset) CalculateRows(int fromRowOffset, int imageHeight, IEnumerable<double> rowHeights)
    {
        double remainingHeightInEmu = fromRowOffset + imageHeight.PixelsToEmu();
        var rowCount = 0;
        var toRowOffset = remainingHeightInEmu;
        const int emuPerRowHeightUnit = 12700;

        foreach (var rowHeight in rowHeights)
        {
            Debug.Assert(rowHeight >= 0, "Row height must be non-negative.");

            remainingHeightInEmu -= rowHeight * emuPerRowHeightUnit;
            if (remainingHeightInEmu < 0)
                return (rowCount, (int)toRowOffset);

            rowCount++;
            toRowOffset = remainingHeightInEmu;
        }

        // TODO: Unreachable?
        return (rowCount, (int)toRowOffset);
    }

    private static bool TryWriteAnchorPart(Span<byte> bytes, ref int bytesWritten, ushort column, uint row, int columnEmuOffset, int rowEmuOffset)
    {
        if (!SpanHelper.TryWrite(column, bytes, ref bytesWritten)) return false;
        if (!"</xdr:col><xdr:colOff>"u8.TryCopyTo(bytes, ref bytesWritten)) return false;
        if (!SpanHelper.TryWrite(columnEmuOffset, bytes, ref bytesWritten)) return false;
        if (!"</xdr:colOff><xdr:row>"u8.TryCopyTo(bytes, ref bytesWritten)) return false;
        if (!SpanHelper.TryWrite(row, bytes, ref bytesWritten)) return false;
        if (!"</xdr:row><xdr:rowOff>"u8.TryCopyTo(bytes, ref bytesWritten)) return false;
        if (!SpanHelper.TryWrite(rowEmuOffset, bytes, ref bytesWritten)) return false;
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
        Anchor,
        ImageEnd,
        Done
    }
}
