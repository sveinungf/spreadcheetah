namespace SpreadCheetah.CellWriters;

internal sealed class RowWriter<T>(
    ICellWriter<T> cellWriter,
    CellWriterState state)
    : BaseRowWriter<T>(cellWriter, state)
    where T : struct
{
    protected override bool TryAddRowCellsForSpan(ReadOnlySpan<T> cells)
    {
        var writerState = State;
        var column = writerState.Column;
        while (column < cells.Length)
        {
            if (!CellWriter.TryWrite(cells[column], writerState))
                return false;

            writerState.Column = ++column;
        }

        return TryWriteRowEnd();
    }
}
