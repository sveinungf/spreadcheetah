namespace SpreadCheetah.CellWriters;

internal sealed class CellWriterState(SpreadsheetBuffer buffer)
{
    public SpreadsheetBuffer Buffer { get; } = buffer;
    public uint NextRowIndex { get; set; } = 1;
    public int Column { get; set; }
}
