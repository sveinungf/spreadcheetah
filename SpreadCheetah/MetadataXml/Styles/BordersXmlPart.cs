using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Diagnostics;
using System.Drawing;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct BordersXmlPart(IList<ImmutableBorder> borders, SpreadsheetBuffer buffer)
{
    private Element _next;
    private int _nextIndex;

    public bool TryWrite()
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
            Element.Header => TryWriteHeader(),
            Element.Borders => TryWriteBorders(),
            _ => buffer.TryWrite("</borders>"u8)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteHeader()
    {
        var span = buffer.GetSpan();
        var written = 0;

        const int defaultCount = 1;
        var totalCount = borders.Count + defaultCount - 1;

        if (!"<borders count=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(totalCount, span, ref written)) return false; // TODO: The count does not include the default border. Need a snapshot test to verify it.

        // The default border must come first
        ReadOnlySpan<byte> defaultBorder = "\">"u8 +
            "<border><left/><right/><top/><bottom/><diagonal/></border>"u8;
        if (!defaultBorder.TryCopyTo(span, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    private bool TryWriteBorders()
    {
        var bordersLocal = borders;

        for (; _nextIndex < bordersLocal.Count; ++_nextIndex)
        {
            var border = bordersLocal[_nextIndex];
            Debug.Assert(border != default);

            var span = buffer.GetSpan();
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

            buffer.Advance(written);
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

        if (color is not { } colorValue)
        {
            if (!"\"/>"u8.TryCopyTo(span, ref written)) return false;
            bytesWritten += written;
            return true;
        }

        if (!"\"><color rgb=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(colorValue, span, ref written)) return false;
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
