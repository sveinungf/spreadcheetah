using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Worksheets;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace SpreadCheetah.CellWriters;

internal class RowWriter<T>(
    ICellWriter<T> cellWriter,
    CellWriterState state)
    where T : struct
{
    private readonly SpreadsheetBuffer Buffer = state.Buffer;

    public bool TryAddRow(IList<T> cells, uint rowIndex)
    {
        if (!CellRowHelper.TryWriteRowStart(rowIndex, Buffer))
        {
            state.Column = -1;
            return false;
        }

        state.Column = 0;
        return TryAddRowCells(cells);
    }

    public bool TryAddRow(ReadOnlySpan<T> cells, uint rowIndex)
    {
        if (!CellRowHelper.TryWriteRowStart(rowIndex, Buffer))
        {
            state.Column = -1;
            return false;
        }

        state.Column = 0;
        return TryAddRowCellsForSpan(cells);
    }

    public bool TryAddRow(IList<T> cells, uint rowIndex, RowOptions options)
    {
        if (!CellRowHelper.TryWriteRowStart(rowIndex, options, Buffer))
        {
            state.Column = -1;
            return false;
        }

        state.Column = 0;
        return options.DefaultStyleId is { } styleId
            ? TryAddRowCells(cells, styleId)
            : TryAddRowCells(cells);
    }

    public bool TryAddRow(ReadOnlySpan<T> cells, uint rowIndex, RowOptions options)
    {
        if (!CellRowHelper.TryWriteRowStart(rowIndex, options, Buffer))
        {
            state.Column = -1;
            return false;
        }

        state.Column = 0;
        return options.DefaultStyleId is { } styleId
            ? TryAddRowCellsForSpan(cells, styleId)
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

    // TODO: Need an implementation of this one that checks for default column styles.
    private bool TryAddRowCellsForSpan(ReadOnlySpan<T> cells)
    {
        var writerState = state;
        var column = writerState.Column;
        while (column < cells.Length)
        {
            if (!cellWriter.TryWrite(cells[column], writerState))
                return false;

            writerState.Column = ++column;
        }

        return TryWriteRowEnd();
    }

    private bool TryAddRowCellsForSpan(ReadOnlySpan<T> cells, StyleId styleId)
    {
        var writerState = state;
        var column = writerState.Column;
        while (column < cells.Length)
        {
            if (!cellWriter.TryWrite(cells[column], styleId, writerState))
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

    private bool TryWriteRowEnd() => Buffer.TryWrite("</row>"u8);

    public async ValueTask AddRowAsync(ReadOnlyMemory<T> cells, uint rowIndex, RowOptions? options, Stream stream, CancellationToken token)
    {
        // If we get here that means that whatever we tried to write didn't fit in the buffer, so just flush right away.
        await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

        EnsureRowStartIsWritten(rowIndex, options);

        var rowStyleId = options?.DefaultStyleId;

        while (state.Column < cells.Length)
        {
            // Attempt to add row cells again
            var beforeIndex = state.Column;
            var success = rowStyleId is null
                ? TryAddRowCellsForSpan(cells.Span)
                : TryAddRowCellsForSpan(cells.Span, rowStyleId);

            if (success)
                return;

            // If no cells were added, the next cell is larger than the buffer.
            if (state.Column == beforeIndex)
            {
                await WriteCellPieceByPieceAsync(cells.Span[state.Column], rowStyleId, stream, token).ConfigureAwait(false);
                ++state.Column;
            }

            // One or more cells were added, repeat.
            await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }

        if (TryWriteRowEnd())
            return;

        await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        TryWriteRowEnd();
    }

    public async ValueTask AddRowAsync(IList<T> cells, uint rowIndex, RowOptions? options, Stream stream, CancellationToken token)
    {
        using var pooledArray = cells.ToPooledArray();
        await AddRowAsync(pooledArray.Memory, rowIndex, options, stream, token).ConfigureAwait(false);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void EnsureRowStartIsWritten(uint rowIndex, RowOptions? options)
    {
        if (state.Column == -1)
        {
            var result = options is null
                ? CellRowHelper.TryWriteRowStart(rowIndex, Buffer)
                : CellRowHelper.TryWriteRowStart(rowIndex, options, Buffer);

            Debug.Assert(result);
            state.Column = 0;
        }
    }

    private async ValueTask WriteCellPieceByPieceAsync(T cell, StyleId? styleId, Stream stream, CancellationToken token)
    {
        // Write start element
        cellWriter.WriteStartElement(cell, styleId, state);

        // Write as much as possible from cell value
        var cellValueIndex = 0;
        while (!cellWriter.TryWriteValue(cell, ref cellValueIndex, state))
        {
            await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }

        // Write end element if it fits in the buffer.
        if (cellWriter.TryWriteEndElement(cell, state))
            return;

        // Flush if the end element doesn't fit. It should always fit after flushing due to the minimum buffer size.
        await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        cellWriter.TryWriteEndElement(cell, state);
    }
}
