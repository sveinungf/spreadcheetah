using SpreadCheetah.CellValueWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling.Internal;
using SpreadCheetah.Worksheets;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace SpreadCheetah.CellWriters;

internal abstract class BaseCellWriter<T>(CellWriterState state, DefaultStyling? defaultStyling)
{
    protected readonly DefaultStyling? DefaultStyling = defaultStyling;
    protected readonly SpreadsheetBuffer Buffer = state.Buffer;
    protected readonly CellWriterState State = state;

    protected abstract bool TryWriteCell(in T cell);
    protected abstract bool WriteStartElement(in T cell);
    protected abstract bool TryWriteEndElement(in T cell);
    protected abstract bool FinishWritingCellValue(in T cell, ref int cellValueIndex);

    public bool TryAddRow(IList<T> cells, uint rowIndex)
    {
        if (!CellRowHelper.TryWriteRowStart(rowIndex, Buffer))
        {
            State.Column = -1;
            return false;
        }

        State.Column = 0;
        return TryAddRowCells(cells);
    }

    public bool TryAddRow(ReadOnlySpan<T> cells, uint rowIndex)
    {
        if (!CellRowHelper.TryWriteRowStart(rowIndex, Buffer))
        {
            State.Column = -1;
            return false;
        }

        State.Column = 0;
        return TryAddRowCellsForSpan(cells);
    }

    public bool TryAddRow(IList<T> cells, uint rowIndex, RowOptions options)
    {
        if (!CellRowHelper.TryWriteRowStart(rowIndex, options, Buffer))
        {
            State.Column = -1;
            return false;
        }

        State.Column = 0;
        return TryAddRowCells(cells);
    }

    public bool TryAddRow(ReadOnlySpan<T> cells, uint rowIndex, RowOptions options)
    {
        if (!CellRowHelper.TryWriteRowStart(rowIndex, options, Buffer))
        {
            State.Column = -1;
            return false;
        }

        State.Column = 0;
        return TryAddRowCellsForSpan(cells);
    }

    private bool TryAddRowCells(IList<T> cells)
    {
        return TryGetSpan(cells, out var span)
            ? TryAddRowCellsForSpan(span)
            : TryAddRowCellsForList(cells);
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

        span = [];
        return false;
    }

    private bool TryAddRowCellsForSpan(ReadOnlySpan<T> cells)
    {
        var writerState = State;
        var column = writerState.Column;
        while (column < cells.Length)
        {
            if (!TryWriteCell(cells[column]))
                return false;

            writerState.Column = ++column;
        }

        return TryWriteRowEnd();
    }

    private bool TryAddRowCellsForList(IList<T> cells)
    {
        var writerState = State;
        var column = writerState.Column;
        while (column < cells.Count)
        {
            if (!TryWriteCell(cells[column]))
                return false;

            writerState.Column = ++column;
        }

        return TryWriteRowEnd();
    }

    private bool TryWriteRowEnd() => Buffer.TryWrite("</row>"u8);

    private async ValueTask WriteRowEndAsync(Stream stream, CancellationToken token)
    {
        if (TryWriteRowEnd())
            return;

        await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        TryWriteRowEnd();
    }

    public async ValueTask AddRowAsync(ReadOnlyMemory<T> cells, uint rowIndex, RowOptions? options, Stream stream, CancellationToken token)
    {
        // If we get here that means that whatever we tried to write didn't fit in the buffer, so just flush right away.
        await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

        EnsureRowStartIsWritten(rowIndex, options);

        while (State.Column < cells.Length)
        {
            // Attempt to add row cells again
            var beforeIndex = State.Column;
            if (TryAddRowCellsForSpan(cells.Span))
                return;

            // If no cells were added, the next cell is larger than the buffer.
            if (State.Column == beforeIndex)
            {
                await WriteCellPieceByPieceAsync(cells.Span[State.Column], stream, token).ConfigureAwait(false);
                ++State.Column;
            }

            // One or more cells were added, repeat.
            await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }

        await WriteRowEndAsync(stream, token).ConfigureAwait(false);
    }

    public async ValueTask AddRowAsync(IList<T> cells, uint rowIndex, RowOptions? options, Stream stream, CancellationToken token)
    {
        // If we get here that means that whatever we tried to write didn't fit in the buffer, so just flush right away.
        await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

        EnsureRowStartIsWritten(rowIndex, options);

        while (State.Column < cells.Count)
        {
            // Attempt to add row cells again
            var beforeIndex = State.Column;
            if (TryAddRowCells(cells))
                return;

            // If no cells were added, the next cell is larger than the buffer.
            if (State.Column == beforeIndex)
            {
                var cell = cells[State.Column];
                await WriteCellPieceByPieceAsync(cell, stream, token).ConfigureAwait(false);
                ++State.Column;
            }

            // One or more cells were added, repeat
            await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        }

        await WriteRowEndAsync(stream, token).ConfigureAwait(false);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void EnsureRowStartIsWritten(uint rowIndex, RowOptions? options)
    {
        if (State.Column == -1)
        {
            var result = options is null
                ? CellRowHelper.TryWriteRowStart(rowIndex, Buffer)
                : CellRowHelper.TryWriteRowStart(rowIndex, options, Buffer);

            Debug.Assert(result);
            State.Column = 0;
        }
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

    protected bool FinishWritingFormulaCellValue(in Cell cell, string formulaText, ref int cellValueIndex)
    {
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);

        // Write the formula
        if (cellValueIndex < formulaText.Length)
        {
            if (!Buffer.WriteLongString(formulaText, ref cellValueIndex)) return false;

            // Finish if there is no cached value to write piece by piece
            if (!writer.CanWriteValuePieceByPiece(cell.DataCell)) return true;
        }

        // If there is a cached value, we need to write "[FORMULA]</f><v>[CACHEDVALUE]"
        var cachedValueStartIndex = formulaText.Length + 1;

        // Write the "</f><v>" part
        if (cellValueIndex < cachedValueStartIndex)
        {
            if (!FormulaCellHelper.EndFormulaBeginCachedValue.TryCopyTo(Buffer.GetSpan()))
                return false;

            Buffer.Advance(FormulaCellHelper.EndFormulaBeginCachedValue.Length);
            cellValueIndex = cachedValueStartIndex;
        }

        // Write the cached value
        var cachedValueIndex = cellValueIndex - cachedValueStartIndex;
        var result = writer.WriteValuePieceByPiece(cell.DataCell, Buffer, ref cachedValueIndex);
        cellValueIndex = cachedValueIndex + cachedValueStartIndex;
        return result;
    }
}
