using SpreadCheetah.Helpers;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.MetadataXml;

internal struct StyleFillsXml
{
    private readonly List<ImmutableFill> _fills;
    private Element _next;
    private int _nextIndex;

    public StyleFillsXml(List<ImmutableFill> fills)
    {
        _fills = fills;
    }

    public bool TryWrite(Span<byte> bytes, ref int bytesWritten)
    {
        if (_next == Element.Header && !Advance(TryWriteHeader(bytes, ref bytesWritten))) return false;
        if (_next == Element.Fills && !Advance(TryWriteFills(bytes, ref bytesWritten))) return false;
        if (_next == Element.Footer && !Advance("</fills>"u8.TryCopyTo(bytes, ref bytesWritten))) return false;

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

        const int defaultCount = 2;
        var totalCount = _fills.Count + defaultCount - 1;

        if (!"<fills count=\""u8.TryCopyTo(span, ref written)) return false;
        if (!SpanHelper.TryWrite(totalCount, span, ref written)) return false;

        // The 2 default fills must come first
        ReadOnlySpan<byte> defaultFills = "\">"u8 +
            """<fill><patternFill patternType="none"/></fill>"""u8 +
            """<fill><patternFill patternType="gray125"/></fill>"""u8;
        if (!defaultFills.TryCopyTo(span, ref written)) return false;

        bytesWritten += written;
        return true;
    }

    private bool TryWriteFills(Span<byte> bytes, ref int bytesWritten)
    {
        var defaultFill = new ImmutableFill();
        var fills = _fills;

        for (; _nextIndex < fills.Count; ++_nextIndex)
        {
            var fill = fills[_nextIndex];
            if (fill.Equals(defaultFill)) continue;
            if (fill.Color is not { } color) continue;

            var span = bytes.Slice(bytesWritten);
            var written = 0;

            if (!"<fill><patternFill patternType=\"solid\"><fgColor rgb=\""u8.TryCopyTo(span, ref written)) return false;
            if (!SpanHelper.TryWrite(color, span, ref written)) return false;
            if (!"\"/></patternFill></fill>"u8.TryCopyTo(span, ref written)) return false;

            bytesWritten += written;
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
