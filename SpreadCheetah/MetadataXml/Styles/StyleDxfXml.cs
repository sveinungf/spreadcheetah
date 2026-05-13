using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Diagnostics;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct StyleDxfXml(
    AddedDifferentialStyle style,
    NumberFormatCounter numberFormatCounter,
    SpreadsheetBuffer buffer)
{
    private NumberFormatXmlPart? _numberFormatPart;
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
        // TODO: Somewhere, we need to check if the style is not just a default style.
        Current = _next switch
        {
            Element.Header => buffer.TryWrite("<dxf>"u8),
            Element.Font => TryWriteFont(),
            Element.NumberFormat => TryWriteNumberFormat(),
            Element.Fill => TryWriteFill(),
            Element.Border => TryWriteBorder(),
            _ => buffer.TryWrite("</dxf>"u8)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteFont()
    {
        var isDefaultFont = style.Font is
        {
            Bold: false,
            Color: null,
            Italic: false,
            Strikethrough: false,
            Underline: Underline.None
        };

        if (isDefaultFont)
            return true;

        Debug.Assert(style.Font is { Name: null, Size: null });
        var xmlPart = new FontXmlPart(buffer, style.Font);
        return xmlPart.TryWrite();
    }

    private bool TryWriteNumberFormat()
    {
        if (style.Format is not { } format)
            return true;

        if (_numberFormatPart is not { } xmlPart)
        {
            var id = numberFormatCounter.GetNextId();
            xmlPart = new NumberFormatXmlPart(id, format);
        }

        if (!xmlPart.TryWrite(buffer))
        {
            _numberFormatPart = xmlPart;
            return false;
        }

        return true;
    }

    private readonly bool TryWriteFill()
    {
        if (style.Fill.Color is not { } color)
            return true;

        return buffer.TryWrite(
            $"{"<fill><patternFill><bgColor rgb=\""u8}" +
            $"{color}" +
            $"{"\"/></patternFill></fill>"u8}");
    }

    private readonly bool TryWriteBorder()
    {
        var isDefaultBorder = style.Border is
        {
            Bottom: null,
            Diagonal: null,
            Left: null,
            Right: null,
            Top: null
        };

        if (isDefaultBorder)
            return true;

        var xmlPart = new BorderXml(style.Border, buffer);
        return xmlPart.TryWrite();
    }

    private enum Element
    {
        Header,
        Font,
        NumberFormat,
        Fill,
        Border,
        Footer,
        Done
    }
}
