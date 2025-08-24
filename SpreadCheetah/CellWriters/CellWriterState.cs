using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellWriters;

internal sealed class CellWriterState(
    SpreadsheetBuffer buffer,
    DefaultStyling? defaultStyling)
{
    public SpreadsheetBuffer Buffer => buffer;
    public DefaultStyling? DefaultStyling => defaultStyling;
    public uint NextRowIndex { get; set; } = 1;
    public int Column { get; set; }
}
