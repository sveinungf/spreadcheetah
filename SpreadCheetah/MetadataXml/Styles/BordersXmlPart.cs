using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Diagnostics;
using System.Drawing;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct BordersXmlPart(IList<ImmutableBorder>? borders, SpreadsheetBuffer buffer)
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
        // The default border must come first
        var count = (borders?.Count ?? 0) + 1;
        return buffer.TryWrite(
            $"{"<borders count=\""u8}" +
            $"{count}" +
            $"{"\"><border><left/><right/><top/><bottom/><diagonal/></border>"u8}");
    }

    private bool TryWriteBorders()
    {
        if (borders is not { } bordersLocal)
            return true;

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

    private static bool TryWriteBorderPartName(BorderPart borderPart, Span<byte> bytes, ref int bytesWritten)
    {
        var name = GetBorderPartName(borderPart);
        return name.TryCopyTo(bytes, ref bytesWritten);
    }

    private static bool TryWriteStyleAttributeValue(BorderStyle style, Span<byte> bytes, ref int bytesWritten)
    {
        var value = GetStyleAttributeValue(style);
        return value.TryCopyTo(bytes, ref bytesWritten);
    }

    private static ReadOnlySpan<byte> GetBorderPartName(BorderPart borderPart) => borderPart switch
    {
        BorderPart.Left => "left"u8,
        BorderPart.Right => "right"u8,
        BorderPart.Top => "top"u8,
        BorderPart.Bottom => "bottom"u8,
        _ => "diagonal"u8
    };

    private static ReadOnlySpan<byte> GetStyleAttributeValue(BorderStyle style) => style switch
    {
        BorderStyle.DashDot => "dashDot"u8,
        BorderStyle.DashDotDot => "dashDotDot"u8,
        BorderStyle.Dashed => "dashed"u8,
        BorderStyle.Dotted => "dotted"u8,
        BorderStyle.DoubleLine => "double"u8,
        BorderStyle.Hair => "hair"u8,
        BorderStyle.Medium => "medium"u8,
        BorderStyle.MediumDashDot => "mediumDashDot"u8,
        BorderStyle.MediumDashDotDot => "mediumDashDotDot"u8,
        BorderStyle.MediumDashed => "mediumDashed"u8,
        BorderStyle.SlantDashDot => "slantDashDot"u8,
        BorderStyle.Thick => "thick"u8,
        _ => "thin"u8
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
