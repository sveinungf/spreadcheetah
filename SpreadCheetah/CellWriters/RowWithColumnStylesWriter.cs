using SpreadCheetah.Styling;

namespace SpreadCheetah.CellWriters;

internal sealed class RowWithColumnStylesWriter<T>(
    ICellWriter<T> cellWriter,
    CellWriterState state,
    IReadOnlyDictionary<int, StyleId> columnStyles)
    : BaseRowWriter<T>(cellWriter, state)
    where T : struct
{
    protected override bool TryAddRowCellsForSpan(ReadOnlySpan<T> cells)
    {
        var writerState = State;
        var column = writerState.Column;
        while (column < cells.Length)
        {
            if (columnStyles.TryGetValue(column, out var styleId))
            {
                if (!CellWriter.TryWrite(cells[column], styleId, writerState))
                    return false;
            }
            else
            {
                if (!CellWriter.TryWrite(cells[column], writerState))
                    return false;
            }

            writerState.Column = ++column;
        }

        return TryWriteRowEnd();
    }

    protected override void WriteStartElement(in T cell, StyleId? rowStyleId)
    {
        var styleId = rowStyleId ?? columnStyles.GetValueOrDefault(State.Column);
        CellWriter.WriteStartElement(cell, styleId, State);
    }
}
