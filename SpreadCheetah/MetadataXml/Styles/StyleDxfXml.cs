using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Diagnostics;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct StyleDxfXml(
    AddedDifferentialStyle style,
    NumberFormatCounter numberFormatCounter,
    SpreadsheetBuffer buffer)
{
    private BorderXml? _borderXmlWriter;
    private FontXmlPart? _fontXmlWriter;
    private NumberFormatXmlPart? _numberFormatXmlWriter;
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

    private bool TryWriteFont()
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

        var xml = _fontXmlWriter ?? new FontXmlPart(buffer, style.Font);
        if (!xml.TryWrite())
        {
            _fontXmlWriter = xml;
            return false;
        }

        return true;
    }

    private bool TryWriteNumberFormat()
    {
        if (style.Format is not { } format)
            return true;

        var xml = _numberFormatXmlWriter
            ?? new NumberFormatXmlPart(numberFormatCounter.GetNextId(), format);

        if (!xml.TryWrite(buffer))
        {
            _numberFormatXmlWriter = xml;
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

    private bool TryWriteBorder()
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

        var xml = _borderXmlWriter ?? new BorderXml(style.Border, buffer);
        if (!xml.TryWrite())
        {
            _borderXmlWriter = xml;
            return false;
        }

        return true;
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
