using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Worksheets;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace SpreadCheetah.CellWriters;

internal abstract class BaseRowWriter<T>(
    ICellWriter<T> cellWriter,
    CellWriterState state)
    where T : struct
{
    private readonly SpreadsheetBuffer _buffer = state.Buffer;

    protected readonly ICellWriter<T> CellWriter = cellWriter;
    protected readonly CellWriterState State = state;

    protected abstract bool TryAddRowCellsForSpan(ReadOnlySpan<T> cells);
    protected abstract void WriteStartElement(in T cell, StyleId? rowStyleId);

    public bool TryAddRow(IList<T> cells, uint rowIndex)
    {
        if (!CellRowHelper.TryWriteRowStart(rowIndex, _buffer))
        {
            State.Column = -1;
            return false;
        }

        State.Column = 0;
        return TryAddRowCells(cells);
    }

    public bool TryAddRow(ReadOnlySpan<T> cells, uint rowIndex)
    {
        if (!CellRowHelper.TryWriteRowStart(rowIndex, _buffer))
        {
            State.Column = -1;
            return false;
        }

        State.Column = 0;
        return TryAddRowCellsForSpan(cells);
    }

    public bool TryAddRow(IList<T> cells, uint rowIndex, RowOptions options, StyleId? rowStyleId)
    {
        if (!CellRowHelper.TryWriteRowStart(rowIndex, options, rowStyleId, _buffer))
        {
            State.Column = -1;
            return false;
        }

        State.Column = 0;
        return rowStyleId is not null
            ? TryAddRowCells(cells, rowStyleId)
            : TryAddRowCells(cells);
    }

    public bool TryAddRow(ReadOnlySpan<T> cells, uint rowIndex, RowOptions options, StyleId? rowStyleId)
    {
        if (!CellRowHelper.TryWriteRowStart(rowIndex, options, rowStyleId, _buffer))
        {
            State.Column = -1;
            return false;
        }

        State.Column = 0;
        return rowStyleId is not null
            ? TryAddRowCellsForSpan(cells, rowStyleId)
            : TryAddRowCellsForSpan(cells);
    }

    private bool TryAddRowCells(IList<T> cells)
    {
        return cells switch
        {
            T[] cellArray => TryAddRowCellsForSpan(cellArray),
#if NET5_0_OR_GREATER
            List<T> cellList => TryAddRowCellsForSpan(System.Runtime.InteropServices.CollectionsMarshal.AsSpan(cellList)),
#endif
            _ => TryAddRowCellsForIList(cells, styleId: null)
        };
    }

    private bool TryAddRowCells(IList<T> cells, StyleId styleId)
    {
        return cells switch
        {
            T[] cellArray => TryAddRowCellsForSpan(cellArray, styleId),
#if NET5_0_OR_GREATER
            List<T> cellList => TryAddRowCellsForSpan(System.Runtime.InteropServices.CollectionsMarshal.AsSpan(cellList), styleId),
#endif
            _ => TryAddRowCellsForIList(cells, styleId)
        };
    }

    private bool TryAddRowCellsForSpan(ReadOnlySpan<T> cells, StyleId styleId)
    {
        var writerState = State;
        var column = writerState.Column;
        while (column < cells.Length)
        {
            if (!CellWriter.TryWrite(cells[column], styleId, writerState))
                return false;

            writerState.Column = ++column;
        }

        return TryWriteRowEnd();
    }

    private bool TryAddRowCellsForIList(IList<T> cells, StyleId? styleId)
    {
        using var pooledArray = cells.ToPooledArray();
        return styleId is null
            ? TryAddRowCellsForSpan(pooledArray.Span)
            : TryAddRowCellsForSpan(pooledArray.Span, styleId);
    }

    protected bool TryWriteRowEnd() => _buffer.TryWrite("</row>"u8);

    public async ValueTask AddRowAsync(ReadOnlyMemory<T> cells, uint rowIndex,
        RowOptions? options, StyleId? rowStyleId, Stream stream, CancellationToken token)
    {
        // If we get here that means that whatever we tried to write didn't fit in the buffer, so just flush right away.
        await _buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

        EnsureRowStartIsWritten(rowIndex, options, rowStyleId);

        while (State.Column < cells.Length)
        {
            // Attempt to add row cells again
            var beforeIndex = State.Column;
            var success = rowStyleId is null
                ? TryAddRowCellsForSpan(cells.Span)
                : TryAddRowCellsForSpan(cells.Span, rowStyleId);

            if (success)
                return;

            // If no cells were added, the next cell is larger than the buffer.
            if (State.Column == beforeIndex)
            {
                await WriteCellPieceByPieceAsync(cells.Span[State.Column], rowStyleId, stream, token).ConfigureAwait(false);
                ++State.Column;
            }

            // One or more cells were added, repeat.
            await _buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }

        if (TryWriteRowEnd())
            return;

        await _buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        TryWriteRowEnd();
    }

    public async ValueTask AddRowAsync(IList<T> cells, uint rowIndex,
        RowOptions? options, StyleId? rowStyleId, Stream stream, CancellationToken token)
    {
        using var pooledArray = cells.ToPooledArray();
        await AddRowAsync(pooledArray.Memory, rowIndex, options, rowStyleId, stream, token).ConfigureAwait(false);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void EnsureRowStartIsWritten(uint rowIndex, RowOptions? options, StyleId? rowStyleId)
    {
        if (State.Column == -1)
        {
            var result = options is null
                ? CellRowHelper.TryWriteRowStart(rowIndex, _buffer)
                : CellRowHelper.TryWriteRowStart(rowIndex, options, rowStyleId, _buffer);

            Debug.Assert(result);
            State.Column = 0;
        }
    }

    private async ValueTask WriteCellPieceByPieceAsync(T cell, StyleId? rowStyleId, Stream stream, CancellationToken token)
    {
        // Write start element
        WriteStartElement(cell, rowStyleId);

        // Write as much as possible from cell value
        var cellValueIndex = 0;
        while (!CellWriter.TryWriteValue(cell, ref cellValueIndex, State))
        {
            await _buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }

        // Write end element if it fits in the buffer.
        if (CellWriter.TryWriteEndElement(cell, State))
            return;

        // Flush if the end element doesn't fit. It should always fit after flushing due to the minimum buffer size.
        await _buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        CellWriter.TryWriteEndElement(cell, State);
    }
}
