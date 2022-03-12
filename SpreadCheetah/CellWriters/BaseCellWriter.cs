using SpreadCheetah.Helpers;
using SpreadCheetah.Worksheets;

namespace SpreadCheetah.CellWriters;

internal abstract class BaseCellWriter<T>
{
    protected readonly SpreadsheetBuffer Buffer;

    protected BaseCellWriter(SpreadsheetBuffer buffer)
    {
        Buffer = buffer;
    }

    protected abstract bool TryWriteCell(in T cell);
    protected abstract bool WriteStartElement(in T cell);
    protected abstract bool TryWriteEndElement(in T cell);
    protected abstract bool FinishWritingCellValue(in T cell, ref int cellValueIndex);

    public bool TryAddRow(IList<T> cells, int rowIndex, out int currentListIndex)
    {
        // Assuming previous actions on the worksheet ensured space in the buffer for row start
        Buffer.Advance(CellRowHelper.GetRowStartBytes(rowIndex, Buffer.GetSpan()));
        return TryAddRowCells(cells, out currentListIndex);
    }

    public bool TryAddRow(ReadOnlySpan<T> cells, int rowIndex, out int currentListIndex)
    {
        // Assuming previous actions on the worksheet ensured space in the buffer for row start
        Buffer.Advance(CellRowHelper.GetRowStartBytes(rowIndex, Buffer.GetSpan()));
        return TryAddRowCellsForSpan(cells, out currentListIndex);
    }

    public bool TryAddRow(IList<T> cells, int rowIndex, RowOptions options, out bool rowStartWritten, out int currentListIndex)
    {
        rowStartWritten = false;
        currentListIndex = 0;

        // Need to check if buffer has enough space. Previous actions only ensure space for a basic row (a row with no options set).
        if (CellRowHelper.ConfiguredRowStartMaxByteCount > Buffer.FreeCapacity)
            return false;

        Buffer.Advance(CellRowHelper.GetRowStartBytes(rowIndex, options, Buffer.GetSpan()));
        rowStartWritten = true;

        return TryAddRowCells(cells, out currentListIndex);
    }

    private bool TryAddRowCells(IList<T> cells, out int currentListIndex) => cells switch
    {
        T[] cellArray => TryAddRowCellsForSpan(cellArray, out currentListIndex),
#if NET5_0_OR_GREATER
        List<T> cellList => TryAddRowCellsForSpan(System.Runtime.InteropServices.CollectionsMarshal.AsSpan(cellList), out currentListIndex),
#endif
        _ => TryAddRowCellsForList(cells, 0, out currentListIndex)
    };

    private bool TryAddRowCells(IList<T> cells, int offset, out int currentListIndex) => cells switch
    {
        T[] cellArray => TryAddRowCellsForSpan(cellArray.AsSpan(offset), out currentListIndex),
#if NET5_0_OR_GREATER
        List<T> cellList => TryAddRowCellsForSpan(System.Runtime.InteropServices.CollectionsMarshal.AsSpan(cellList).Slice(offset), out currentListIndex),
#endif
        _ => TryAddRowCellsForList(cells, offset, out currentListIndex)
    };

    private bool TryAddRowCellsForSpan(ReadOnlySpan<T> cells, out int currentListIndex)
    {
        for (currentListIndex = 0; currentListIndex < cells.Length; ++currentListIndex)
        {
            // Write cell if it fits in the buffer
            if (!TryWriteCell(cells[currentListIndex]))
                return false;
        }

        return WriteRowEnd();
    }

    private bool TryAddRowCellsForList(IList<T> cells, int offset, out int currentListIndex)
    {
        for (currentListIndex = offset; currentListIndex < cells.Count; ++currentListIndex)
        {
            // Write cell if it fits in the buffer
            if (!TryWriteCell(cells[currentListIndex]))
                return false;
        }

        return WriteRowEnd();
    }

    // Also ensuring space in the buffer for starting another basic row, so that we might not need to check space in the buffer twice.
    private bool CanWriteRowEnd() => Buffer.FreeCapacity>= CellRowHelper.RowEnd.Length + CellRowHelper.BasicRowStartMaxByteCount;

    private void DoWriteRowEnd() => Buffer.Advance(SpanHelper.GetBytes(CellRowHelper.RowEnd, Buffer.GetSpan()));

    private bool WriteRowEnd()
    {
        if (!CanWriteRowEnd())
            return false;

        DoWriteRowEnd();
        return true;
    }

    private async ValueTask WriteRowEndAsync(Stream stream, CancellationToken token)
    {
        if (!CanWriteRowEnd())
            await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

        DoWriteRowEnd();
    }

    public async ValueTask AddRowAsync(IList<T> cells, int currentIndex, Stream stream, CancellationToken token)
    {
        while (currentIndex < cells.Count)
        {
            // If we get here that means that the next cell didn't fit in the buffer, so just flush right away
            await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

            // Attempt to add row cells again
            var beforeIndex = currentIndex;
            if (TryAddRowCells(cells, currentIndex, out currentIndex))
                return;

            // One or more cells were added, repeat
            if (currentIndex != beforeIndex)
                continue;

            // No cells were added, which means the next cell is larger than the buffer
            var cell = cells[currentIndex];
            await WriteCellPieceByPieceAsync(cell, stream, token).ConfigureAwait(false);
            ++currentIndex;
        }

        await WriteRowEndAsync(stream, token).ConfigureAwait(false);
    }

    public async ValueTask AddRowAsync(ReadOnlyMemory<T> cells, Stream stream, CancellationToken token)
    {
        while (!cells.IsEmpty)
        {
            // If we get here that means that the next cell didn't fit in the buffer, so just flush right away.
            await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

            // Attempt to add row cells again
            if (TryAddRowCellsForSpan(cells.Span, out var currentIndex))
                return;

            // If no cells were added, the next cell is larger than the buffer.
            if (currentIndex == 0)
            {
                await WriteCellPieceByPieceAsync(cells.Span[0], stream, token).ConfigureAwait(false);
                currentIndex = 1;
            }

            // One or more cells were added, repeat.
            cells = cells.Slice(currentIndex);
        }

        await WriteRowEndAsync(stream, token).ConfigureAwait(false);
    }

    public async ValueTask AddRowAsync(IList<T> cells, int rowIndex, RowOptions options, bool rowStartWritten, int currentCellIndex, int endCellIndex, Stream stream, CancellationToken token)
    {
        if (!rowStartWritten)
        {
            // If we get here that means that whatever we tried to write didn't fit in the buffer, so just flush right away.
            await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
            Buffer.Advance(CellRowHelper.GetRowStartBytes(rowIndex, options, Buffer.GetSpan()));
        }

        await AddRowAsync(cells, currentCellIndex, stream, token).ConfigureAwait(false);
    }

    private async ValueTask WriteCellPieceByPieceAsync(T cell, Stream stream, CancellationToken token)
    {
        // Write start element
        WriteStartElement(cell);

        // Write as much as possible from cell value
        var cellValueIndex = 0;
        while (!FinishWritingCellValue(cell, ref cellValueIndex))
        {
            await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }

        // Write end element if it fits in the buffer.
        if (TryWriteEndElement(cell))
            return;

        // Flush if the end element doesn't fit. It should always fit after flushing due to the minimum buffer size.
        await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        TryWriteEndElement(cell);
    }
}
