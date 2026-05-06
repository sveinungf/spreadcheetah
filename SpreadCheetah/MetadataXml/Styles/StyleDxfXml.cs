using SpreadCheetah.Styling.Internal;

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
        // TODO: Handle more parts of the style.
        // TODO: Somewhere, we need to check if the style is not just a default style.
        Current = _next switch
        {
            Element.Header => buffer.TryWrite("<dxf>"u8),
            Element.NumberFormat => TryWriteNumberFormat(),
            Element.Fill => TryWriteFill(),
            _ => buffer.TryWrite("</dxf>"u8)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
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

    private enum Element
    {
        Header,
        NumberFormat,
        Fill,
        Footer,
        Done
    }
}
