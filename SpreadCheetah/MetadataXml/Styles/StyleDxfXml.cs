using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct StyleDxfXml(
    AddedDifferentialStyle style,
    SpreadsheetBuffer buffer)
{
    private Element _next;

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
        // TODO: Handle more parts of the style.
        // TODO: Somewhere, we need to check if the style is not just a default style.
        Current = _next switch
        {
            Element.Header => buffer.TryWrite("<dxf>"u8),
            Element.Fill => TryWriteFill(),
            _ => buffer.TryWrite("</dxf>"u8)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
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
        Fill,
        Footer,
        Done
    }
}
