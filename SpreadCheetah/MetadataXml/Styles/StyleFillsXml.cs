using SpreadCheetah.Helpers;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct StyleFillsXml(List<ImmutableFill> fills, SpreadsheetBuffer buffer)
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
            Element.Fills => TryWriteFills(),
            _ => buffer.TryWrite("</fills>"u8)
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteHeader()
    {
        var span = buffer.GetSpan();
        var written = 0;

        const int defaultCount = 2;
        var totalCount = fills.Count + defaultCount - 1;

        if (!"<fills count=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(totalCount, span, ref written)) return false;

        // The 2 default fills must come first
        ReadOnlySpan<byte> defaultFills = "\">"u8 +
            """<fill><patternFill patternType="none"/></fill>"""u8 +
            """<fill><patternFill patternType="gray125"/></fill>"""u8;
        if (!defaultFills.TryCopyTo(span, ref written)) return false;

        buffer.Advance(written);
        return true;
    }

    private bool TryWriteFills()
    {
        var fillsLocal = fills;

        for (; _nextIndex < fillsLocal.Count; ++_nextIndex)
        {
            var fill = fillsLocal[_nextIndex];
            if (fill.Equals(default)) continue;
            if (fill.Color is not { } color) continue;

            var span = buffer.GetSpan();
            var written = 0;

            if (!"<fill><patternFill patternType=\"solid\"><fgColor rgb=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(color, span, ref written)) return false;
            if (!"\"/></patternFill></fill>"u8.TryCopyTo(span, ref written)) return false;

            buffer.Advance(written);
        }

        return true;
    }

    private enum Element
    {
        Header,
        Fills,
        Footer,
        Done
    }
}
