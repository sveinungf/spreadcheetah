using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Drawing;

namespace SpreadCheetah.MetadataXml;

internal struct StyleBordersXml
{
    private readonly List<ImmutableBorder> _borders;
    private Element _next;
    private int _nextIndex;

    public StyleBordersXml(List<ImmutableBorder> borders)
    {
        _borders = borders;
    }

    public bool TryWrite(Span<byte> bytes, ref int bytesWritten)
    {
        if (_next == Element.Header && !Advance(TryWriteHeader(bytes, ref bytesWritten))) return false;
        if (_next == Element.Borders && !Advance(TryWriteBorders(bytes, ref bytesWritten))) return false;
        if (_next == Element.Footer && !Advance("</borders>"u8.TryCopyTo(bytes, ref bytesWritten))) return false;

        return true;
    }

    private bool Advance(bool success)
    {
        if (success)
            ++_next;

        return success;
    }

    private readonly bool TryWriteHeader(Span<byte> bytes, ref int bytesWritten)
    {
        var span = bytes.Slice(bytesWritten);
        var written = 0;

        const int defaultCount = 1;
        var totalCount = _borders.Count + defaultCount - 1;

        if (!"<borders count=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(totalCount, span, ref written)) return false;

        // The default border must come first
        ReadOnlySpan<byte> defaultBorder = "\">"u8 +
            "<border><left/><right/><top/><bottom/><diagonal/></border>"u8;
        if (!defaultBorder.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private bool TryWriteBorders(Span<byte> bytes, ref int bytesWritten)
    {
        var defaultBorder = new ImmutableBorder();
        var borders = _borders;

        for (; _nextIndex < borders.Count; ++_nextIndex)
        {
            var border = borders[_nextIndex];
            if (border.Equals(defaultBorder)) continue;

            var span = bytes.Slice(bytesWritten);
            var written = 0;

            var diag = border.Diagonal;
            if (!"<border"u8.TryCopyTo(span, ref written)) return false;
            if (diag.Type.HasFlag(DiagonalBorderType.DiagonalUp) && !" diagonalUp=\"true\""u8.TryCopyTo(span, ref written)) return false;
            if (diag.Type.HasFlag(DiagonalBorderType.DiagonalDown) && !" diagonalDown=\"true\""u8.TryCopyTo(span, ref written)) return false;
            if (!">"u8.TryCopyTo(span, ref written)) return false;

            if (!TryWriteBorderPart(BorderPart.Left, border.Left.BorderStyle, border.Left.Color, span, ref written)) return false;
            if (!TryWriteBorderPart(BorderPart.Right, border.Right.BorderStyle, border.Right.Color, span, ref written)) return false;
            if (!TryWriteBorderPart(BorderPart.Top, border.Top.BorderStyle, border.Top.Color, span, ref written)) return false;
            if (!TryWriteBorderPart(BorderPart.Bottom, border.Bottom.BorderStyle, border.Bottom.Color, span, ref written)) return false;
            if (!TryWriteBorderPart(BorderPart.Diagonal, border.Diagonal.BorderStyle, border.Diagonal.Color, span, ref written)) return false;

            if (!"</border>"u8.TryCopyTo(span, ref written)) return false;

            bytesWritten += written;
        }

        return true;
    }

    private static bool TryWriteBorderPart(BorderPart borderPart, BorderStyle style, Color? color, Span<byte> bytes, ref int bytesWritten)
    {
        var span = bytes.Slice(bytesWritten);
        var written = 0;

        if (!"<"u8.TryCopyTo(span, ref written)) return false;
        if (!TryWriteBorderPartName(borderPart, span, ref written)) return false;

        if (style == BorderStyle.None)
        {
            if (!"/>"u8.TryCopyTo(span, ref written)) return false;
            bytesWritten += written;
            return true;
        }

        if (!" style=\""u8.TryCopyTo(span, ref written)) return false;
        if (!TryWriteStyleAttributeValue(style, span, ref written)) return false;

        if (color is null)
        {
            if (!"\"/>"u8.TryCopyTo(span, ref written)) return false;
            bytesWritten += written;
            return true;
        }

        if (!"\"><color rgb=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(color.Value, span, ref written)) return false;
        if (!"\"/></"u8.TryCopyTo(span, ref written)) return false;
        if (!TryWriteBorderPartName(borderPart, span, ref written)) return false;
        if (!">"u8.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private static bool TryWriteBorderPartName(BorderPart borderPart, Span<byte> bytes, ref int bytesWritten) => borderPart switch
    {
        BorderPart.Left => "left"u8.TryCopyTo(bytes, ref bytesWritten),
        BorderPart.Right => "right"u8.TryCopyTo(bytes, ref bytesWritten),
        BorderPart.Top => "top"u8.TryCopyTo(bytes, ref bytesWritten),
        BorderPart.Bottom => "bottom"u8.TryCopyTo(bytes, ref bytesWritten),
        BorderPart.Diagonal => "diagonal"u8.TryCopyTo(bytes, ref bytesWritten),
        _ => true
    };

    private static bool TryWriteStyleAttributeValue(BorderStyle style, Span<byte> bytes, ref int bytesWritten) => style switch
    {
        BorderStyle.DashDot => "dashDot"u8.TryCopyTo(bytes, ref bytesWritten),
        BorderStyle.DashDotDot => "dashDotDot"u8.TryCopyTo(bytes, ref bytesWritten),
        BorderStyle.Dashed => "dashed"u8.TryCopyTo(bytes, ref bytesWritten),
        BorderStyle.Dotted => "dotted"u8.TryCopyTo(bytes, ref bytesWritten),
        BorderStyle.DoubleLine => "double"u8.TryCopyTo(bytes, ref bytesWritten),
        BorderStyle.Hair => "hair"u8.TryCopyTo(bytes, ref bytesWritten),
        BorderStyle.Medium => "medium"u8.TryCopyTo(bytes, ref bytesWritten),
        BorderStyle.MediumDashDot => "mediumDashDot"u8.TryCopyTo(bytes, ref bytesWritten),
        BorderStyle.MediumDashDotDot => "mediumDashDotDot"u8.TryCopyTo(bytes, ref bytesWritten),
        BorderStyle.MediumDashed => "mediumDashed"u8.TryCopyTo(bytes, ref bytesWritten),
        BorderStyle.SlantDashDot => "slantDashDot"u8.TryCopyTo(bytes, ref bytesWritten),
        BorderStyle.Thick => "thick"u8.TryCopyTo(bytes, ref bytesWritten),
        BorderStyle.Thin => "thin"u8.TryCopyTo(bytes, ref bytesWritten),
        _ => true
    };

    private enum BorderPart
    {
        Left,
        Right,
        Top,
        Bottom,
        Diagonal
    }

    private enum Element
    {
        Header,
        Borders,
        Footer,
        Done
    }
}
