using SpreadCheetah.Styling.Internal;
using System.Diagnostics;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct FillsXmlPart(IList<ImmutableFill>? fills, SpreadsheetBuffer buffer)
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
        // The 2 default fills must come first
        var count = (fills?.Count ?? 0) + 2;
        var defaultFills = "\">"u8 +
            """<fill><patternFill patternType="none"/></fill>"""u8 +
            """<fill><patternFill patternType="gray125"/></fill>"""u8;

        return buffer.TryWrite(
            $"{"<fills count=\""u8}" +
            $"{count}" +
            $"{defaultFills}");
    }

    private bool TryWriteFills()
    {
        if (fills is not { } fillsLocal)
            return true;

        for (; _nextIndex < fillsLocal.Count; ++_nextIndex)
        {
            var fill = fillsLocal[_nextIndex];
            Debug.Assert(fill != default);
            if (fill.Color is not { } color) continue;

            var success = buffer.TryWrite(
                $"{"<fill><patternFill patternType=\"solid\"><fgColor rgb=\""u8}" +
                $"{color}" +
                $"{"\"/></patternFill></fill>"u8}");

            if (!success)
                return false;
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
