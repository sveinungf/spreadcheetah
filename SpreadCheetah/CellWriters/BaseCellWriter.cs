using SpreadCheetah.Helpers;
using SpreadCheetah.Styling.Internal;
using SpreadCheetah.Worksheets;

namespace SpreadCheetah.CellWriters;

internal abstract class BaseCellWriter<T>
{
    protected readonly DefaultStyling? DefaultStyling;
    protected readonly SpreadsheetBuffer Buffer;
    protected readonly CellWriterState State;

    protected BaseCellWriter(CellWriterState state, DefaultStyling? defaultStyling)
    {
        Buffer = state.Buffer;
        DefaultStyling = defaultStyling;
        State = state;
    }

    protected abstract bool TryWriteCell(in T cell);
    protected abstract bool WriteStartElement(in T cell);
    protected abstract bool TryWriteEndElement(in T cell);
    protected abstract bool FinishWritingCellValue(in T cell, ref int cellValueIndex);

    public bool TryAddRow(IList<T> cells, uint rowIndex, out int currentListIndex)
    {
        currentListIndex = -1;
        return CellRowHelper.TryWriteRowStart(rowIndex, Buffer) && TryAddRowCells(cells, out currentListIndex);
    }

    public bool TryAddRow(ReadOnlySpan<T> cells, uint rowIndex, out int currentListIndex)
    {
        currentListIndex = -1;
        return CellRowHelper.TryWriteRowStart(rowIndex, Buffer) && TryAddRowCellsForSpan(cells, out currentListIndex);
    }

    public bool TryAddRow(IList<T> cells, uint rowIndex, RowOptions options, out int currentListIndex)
    {
        currentListIndex = -1;
        return CellRowHelper.TryWriteRowStart(rowIndex, options, Buffer) && TryAddRowCells(cells, out currentListIndex);
    }

    public bool TryAddRow(ReadOnlySpan<T> cells, uint rowIndex, RowOptions options, out int currentListIndex)
    {
        currentListIndex = -1;
        return CellRowHelper.TryWriteRowStart(rowIndex, options, Buffer) && TryAddRowCellsForSpan(cells, out currentListIndex);
    }

    private bool TryAddRowCells(IList<T> cells, out int currentListIndex)
    {
        return TryGetSpan(cells, out var span)
            ? TryAddRowCellsForSpan(span, out currentListIndex)
            : TryAddRowCellsForList(cells, 0, out currentListIndex);
    }

    private bool TryAddRowCells(IList<T> cells, int offset, out int currentListIndex)
    {
        if (TryGetSpan(cells, out var span))
        {
            var result = TryAddRowCellsForSpan(span.Slice(offset), out var spanIndex);
            currentListIndex = offset + spanIndex;
            return result;
        }

        return TryAddRowCellsForList(cells, offset, out currentListIndex);
    }

    private static bool TryGetSpan(IList<T> cells, out ReadOnlySpan<T> span)
    {
        if (cells is T[] cellArray)
        {
            span = cellArray.AsSpan();
            return true;
        }

#if NET5_0_OR_GREATER
        if (cells is List<T> cellList)
        {
            span = System.Runtime.InteropServices.CollectionsMarshal.AsSpan(cellList);
            return true;
        }
#endif

        span = ReadOnlySpan<T>.Empty;
        return false;
    }

    private bool TryAddRowCellsForSpan(ReadOnlySpan<T> cells, out int spanIndex)
    {
        for (spanIndex = 0; spanIndex < cells.Length; ++spanIndex)
        {
            // Write cell if it fits in the buffer
            if (!TryWriteCell(cells[spanIndex]))
                return false;
        }

        return TryWriteRowEnd();
    }

    private bool TryAddRowCellsForList(IList<T> cells, int offset, out int currentListIndex)
    {
        for (currentListIndex = offset; currentListIndex < cells.Count; ++currentListIndex)
        {
            // Write cell if it fits in the buffer
            if (!TryWriteCell(cells[currentListIndex]))
                return false;
        }

        return TryWriteRowEnd();
    }

    private bool TryWriteRowEnd()
    {
        var rowEnd = "</row>"u8;
        if (rowEnd.TryCopyTo(Buffer.GetSpan()))
        {
            Buffer.Advance(rowEnd.Length);
            return true;
        }

        return false;
    }

    private async ValueTask WriteRowEndAsync(Stream stream, CancellationToken token)
    {
        if (TryWriteRowEnd())
            return;

        await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        TryWriteRowEnd();
    }

    public async ValueTask AddRowAsync(ReadOnlyMemory<T> cells, uint rowIndex, int currentCellIndex, Stream stream, CancellationToken token)
    {
        // If we get here that means that the next cell didn't fit in the buffer, so just flush right away.
        await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

        if (currentCellIndex == -1)
        {
            CellRowHelper.TryWriteRowStart(rowIndex, Buffer);
            currentCellIndex = 0;
        }

        await AddRowCellsAsync(cells.Slice(currentCellIndex), stream, token).ConfigureAwait(false);
    }

    public async ValueTask AddRowAsync(IList<T> cells, uint rowIndex, int currentCellIndex, Stream stream, CancellationToken token)
    {
        // If we get here that means that the next cell didn't fit in the buffer, so just flush right away.
        await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

        if (currentCellIndex == -1)
        {
            CellRowHelper.TryWriteRowStart(rowIndex, Buffer);
            currentCellIndex = 0;
        }

        await AddRowCellsAsync(cells, currentCellIndex, stream, token).ConfigureAwait(false);
    }

    public async ValueTask AddRowAsync(ReadOnlyMemory<T> cells, uint rowIndex, RowOptions options, int currentCellIndex, Stream stream, CancellationToken token)
    {
        // If we get here that means that whatever we tried to write didn't fit in the buffer, so just flush right away.
        await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

        if (currentCellIndex == -1)
        {
            CellRowHelper.TryWriteRowStart(rowIndex, options, Buffer);
            currentCellIndex = 0;
        }

        await AddRowCellsAsync(cells.Slice(currentCellIndex), stream, token).ConfigureAwait(false);
    }

    public async ValueTask AddRowAsync(IList<T> cells, uint rowIndex, RowOptions options, int currentCellIndex, Stream stream, CancellationToken token)
    {
        // If we get here that means that whatever we tried to write didn't fit in the buffer, so just flush right away.
        await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

        if (currentCellIndex == -1)
        {
            CellRowHelper.TryWriteRowStart(rowIndex, options, Buffer);
            currentCellIndex = 0;
        }

        await AddRowCellsAsync(cells, currentCellIndex, stream, token).ConfigureAwait(false);
    }

    private async ValueTask AddRowCellsAsync(ReadOnlyMemory<T> cells, Stream stream, CancellationToken token)
    {
        while (!cells.IsEmpty)
        {
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
            await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }

        await WriteRowEndAsync(stream, token).ConfigureAwait(false);
    }

    private async ValueTask AddRowCellsAsync(IList<T> cells, int currentIndex, Stream stream, CancellationToken token)
    {
        while (currentIndex < cells.Count)
        {
            // Attempt to add row cells again
            var beforeIndex = currentIndex;
            if (TryAddRowCells(cells, currentIndex, out currentIndex))
                return;

            // If no cells were added, the next cell is larger than the buffer.
            if (currentIndex == beforeIndex)
            {
                var cell = cells[currentIndex];
                await WriteCellPieceByPieceAsync(cell, stream, token).ConfigureAwait(false);
                ++currentIndex;
            }

            // One or more cells were added, repeat
            await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }

        await WriteRowEndAsync(stream, token).ConfigureAwait(false);
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
