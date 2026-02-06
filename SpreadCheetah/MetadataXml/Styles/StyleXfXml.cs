using SpreadCheetah.MetadataXml.Attributes;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.MetadataXml.Styles;

internal struct StyleXfXml(
    AddedStyle addedStyle,
    int? embeddedNamedStyleIndex,
    IList<ImmutableAlignment>? alignments,
    bool cellXfsEntry,
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
            Element.NumberFormat => TryWriteNumberFormat(),
            Element.Font => TryWriteFont(),
            Element.Fill => TryWriteFill(),
            Element.Border => TryWriteBorder(),
            Element.XfId => TryWriteXfId(),
            Element.Alignment => TryWriteAlignment(),
            _ => TryWriteXfEnd()
        };

        if (Current)
            ++_next;

        return _next < Element.Done;
    }

    private readonly bool TryWriteNumberFormat()
    {
        var numberFormatId = GetNumberFormatId();
        var applyNumberFormat = new BooleanAttribute("applyNumberFormat"u8, numberFormatId != 0 ? true : null);

        return buffer.TryWrite(
            $"{"<xf numFmtId=\""u8}" +
            $"{numberFormatId}" +
            $"{"\""u8}" +
            $"{applyNumberFormat}");
    }

    private readonly int GetNumberFormatId()
    {
        const int customFormatStartId = 165;

        return addedStyle switch
        {
            { StandardFormat: { } standardFormat } => (int)standardFormat,
            { CustomFormatIndex: { } customIndex } => customIndex + customFormatStartId,
            _ => 0
        };
    }

    private readonly bool TryWriteFont()
    {
        const int defaultFonts = 1;
        var fontIndex = (addedStyle.FontIndex + defaultFonts) ?? 0;
        var applyFont = new BooleanAttribute("applyFont"u8, fontIndex != 0 ? true : null);

        return buffer.TryWrite(
            $"{" fontId=\""u8}" +
            $"{fontIndex}" +
            $"{"\""u8}" +
            $"{applyFont}");
    }

    private readonly bool TryWriteFill()
    {
        const int defaultFills = 2;
        var fillIndex = (addedStyle.FillIndex + defaultFills) ?? 0;
        var applyFill = new BooleanAttribute("applyFill"u8, fillIndex != 0 ? true : null);

        return buffer.TryWrite(
            $"{" fillId=\""u8}" +
            $"{fillIndex}" +
            $"{"\""u8}" +
            $"{applyFill}");
    }

    private readonly bool TryWriteBorder()
    {
        if (addedStyle.BorderIndex is not { } borderIndex)
            return true;

        const int defaultBorders = 1;
        return buffer.TryWrite(
            $"{" borderId=\""u8}" +
            $"{borderIndex + defaultBorders}" +
            $"{"\" applyBorder=\"1\""u8}");
    }

    private readonly bool TryWriteXfId()
    {
        if (!cellXfsEntry)
            return true;

        var xfId = (embeddedNamedStyleIndex + 1) ?? 0;
        return buffer.TryWrite($"{" xfId=\""u8}{xfId}{"\""u8}");
    }

    private readonly bool TryWriteAlignment()
    {
        if (addedStyle.AlignmentIndex is not { } index || alignments is null)
            return true;

        var alignment = alignments[index];
        var horizontalAlignment = GetHorizontalAlignmentValue(alignment.Horizontal);
        var verticalAlignment = GetVerticalAlignmentValue(alignment.Vertical);
        var wrapText = new BooleanAttribute("wrapText"u8, alignment.WrapText ? true : null);
        var indent = new IntAttribute("indent"u8, alignment.Indent > 0 ? alignment.Indent : null);

        return buffer.TryWrite(
            $"{""" applyAlignment="1"><alignment"""u8}" +
            $"{horizontalAlignment}" +
            $"{verticalAlignment}" +
            $"{wrapText}" +
            $"{indent}" +
            $"{"/>"u8}");
    }

    private readonly bool TryWriteXfEnd()
    {
        var end = addedStyle.AlignmentIndex is null || alignments is null
            ? "/>"u8
            : "</xf>"u8;

        return buffer.TryWrite(end);
    }

    private static ReadOnlySpan<byte> GetHorizontalAlignmentValue(HorizontalAlignment alignment) => alignment switch
    {
        HorizontalAlignment.Left => " horizontal=\"left\""u8,
        HorizontalAlignment.Center => " horizontal=\"center\""u8,
        HorizontalAlignment.Right => " horizontal=\"right\""u8,
        _ => []
    };

    private static ReadOnlySpan<byte> GetVerticalAlignmentValue(VerticalAlignment alignment) => alignment switch
    {
        VerticalAlignment.Center => " vertical=\"center\""u8,
        VerticalAlignment.Top => " vertical=\"top\""u8,
        _ => []
    };

    private enum Element
    {
        NumberFormat,
        Font,
        Fill,
        Border,
        XfId,
        Alignment,
        XfEnd,
        Done
    }
}
