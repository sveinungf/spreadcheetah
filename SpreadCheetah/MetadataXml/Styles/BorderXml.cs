using SpreadCheetah.MetadataXml.Attributes;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Drawing;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct BorderXml(
    ImmutableBorder border,
    SpreadsheetBuffer buffer)
{
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
            Element.BorderStart => TryWriteBorderStart(),
            Element.Left => TryWriteBorderPart("left"u8, border.Left.BorderStyle, border.Left.Color),
            Element.Right => TryWriteBorderPart("right"u8, border.Right.BorderStyle, border.Right.Color),
            Element.Top => TryWriteBorderPart("top"u8, border.Top.BorderStyle, border.Top.Color),
            Element.Bottom => TryWriteBorderPart("bottom"u8, border.Bottom.BorderStyle, border.Bottom.Color),
            Element.Diagonal => TryWriteBorderPart("diagonal"u8, border.Diagonal.BorderStyle, border.Diagonal.Color),
            _ => buffer.TryWrite("</border>"u8),
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteBorderStart()
    {
        var hasDiagonalUp = border.Diagonal.Type.HasFlag(DiagonalBorderType.DiagonalUp);
        var diagonalUp = new SpanByteAttribute("diagonalUp"u8, hasDiagonalUp ? "true"u8 : []);

        var hasDiagonalDown = border.Diagonal.Type.HasFlag(DiagonalBorderType.DiagonalDown);
        var diagonalDown = new SpanByteAttribute("diagonalDown"u8, hasDiagonalDown ? "true"u8 : []);

        return buffer.TryWrite($"{"<border"u8}{diagonalUp}{diagonalDown}{">"u8}");
    }

    private readonly bool TryWriteBorderPart(
        ReadOnlySpan<byte> borderPart,
        BorderStyle style,
        Color? color)
    {
        if (style == BorderStyle.None)
            return buffer.TryWrite($"{"<"u8}{borderPart}{"/>"u8}");

        var styleValue = GetStyleValue(style);

        if (color is not { } colorValue)
        {
            return buffer.TryWrite(
                $"{"<"u8}" +
                $"{borderPart}" +
                $"{" style=\""u8}" +
                $"{styleValue}" +
                $"{"\"/>"u8}");
        }

        return buffer.TryWrite(
            $"{"<"u8}" +
            $"{borderPart}" +
            $"{" style=\""u8}" +
            $"{styleValue}" +
            $"{"\"><color rgb=\""u8}" +
            $"{colorValue}" +
            $"{"\"/></"u8}" +
            $"{borderPart}" +
            $"{">"u8}");
    }

    private static ReadOnlySpan<byte> GetStyleValue(BorderStyle style) => style switch
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

    private enum Element
    {
        BorderStart,
        Left,
        Right,
        Top,
        Bottom,
        Diagonal,
        BorderEnd,
        Done
    }
}
