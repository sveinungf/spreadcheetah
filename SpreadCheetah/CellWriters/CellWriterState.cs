using SpreadCheetah.Helpers;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellWriters;

internal sealed class CellWriterState(
    SpreadsheetBuffer buffer,
    DefaultStyling? defaultStyling,
    InlineXmlTags inlineXmlTags)
{
    public SpreadsheetBuffer Buffer => buffer;
    public DefaultStyling? DefaultStyling => defaultStyling;
    public InlineXmlTags InlineXmlTags => inlineXmlTags;
    public uint NextRowIndex { get; set; } = 1;
    public int Column { get; set; }
}
