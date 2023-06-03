namespace SpreadCheetah.CellWriters;

internal sealed class CellWriterState
{
    public SpreadsheetBuffer Buffer { get; }
    public uint NextRowIndex { get; set; } = 1;
    public int Column { get; set; }

    public CellWriterState(SpreadsheetBuffer buffer)
    {
        Buffer = buffer;
    }
}
