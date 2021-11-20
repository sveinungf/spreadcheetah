using SpreadCheetah.Helpers;

namespace SpreadCheetah.CellWriters
{
    internal abstract class BaseCellWriter<T>
    {
        protected readonly SpreadsheetBuffer Buffer;

        protected BaseCellWriter(SpreadsheetBuffer buffer)
        {
            Buffer = buffer;
        }

        protected abstract bool TryWriteCell(in T cell, out int bytesNeeded);
        protected abstract bool GetBytes(in T cell, bool assertSize);
        protected abstract bool WriteStartElement(in T cell);
        protected abstract bool TryWriteEndElement(in T cell);
        protected abstract bool FinishWritingCellValue(in T cell, ref int cellValueIndex);

        public bool TryAddRow(IList<T> cells, int rowIndex, out int currentListIndex)
        {
            // Assuming previous actions on the worksheet ensured space in the buffer for row start
            Buffer.Advance(GetRowStartBytes(rowIndex, Buffer.GetNextSpan()));

            for (currentListIndex = 0; currentListIndex < cells.Count; ++currentListIndex)
            {
                // Write cell if it fits in the buffer
                if (!TryWriteCell(cells[currentListIndex], out _))
                    return false;
            }

            // Also ensuring space in the buffer for the next row start, so that we don't need to check space in the buffer twice
            if (CellWriterHelper.RowEnd.Length + CellWriterHelper.RowStartMaxByteCount > Buffer.GetRemainingBuffer())
                return false;

            Buffer.Advance(SpanHelper.GetBytes(CellWriterHelper.RowEnd, Buffer.GetNextSpan()));
            return true;
        }

        private static int GetRowStartBytes(int rowIndex, Span<byte> bytes)
        {
            var bytesWritten = SpanHelper.GetBytes(CellWriterHelper.RowStart, bytes);
            bytesWritten += Utf8Helper.GetBytes(rowIndex, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(CellWriterHelper.RowStartEndTag, bytes.Slice(bytesWritten));
            return bytesWritten;
        }

        public async ValueTask AddRowAsync(IList<T> cells, int currentIndex, Stream stream, CancellationToken token)
        {
            // If we get here that means that the next cell didn't fit in the buffer, so just flush right away
            await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

            for (var i = currentIndex; i < cells.Count; ++i)
            {
                var cell = cells[i];

                // Write cell if it fits in the buffer
                if (TryWriteCell(cell, out var bytesNeeded))
                    continue;

                await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

                // Write cell if it fits in the buffer
                if (bytesNeeded <= Buffer.GetRemainingBuffer())
                {
                    GetBytes(cell, false);
                    continue;
                }

                await WriteCellPieceByPieceAsync(cell, stream, token).ConfigureAwait(false);
            }

            // Also ensuring space in the buffer for the next row start, so that we don't need to check space in the buffer twice
            if (CellWriterHelper.RowEnd.Length + CellWriterHelper.RowStartMaxByteCount > Buffer.GetRemainingBuffer())
                await Buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);

            Buffer.Advance(SpanHelper.GetBytes(CellWriterHelper.RowEnd, Buffer.GetNextSpan()));
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

        protected bool FinishWritingCellValue(string cellValue, ref int cellValueIndex)
        {
            var remainingBuffer = Buffer.GetRemainingBuffer();
            var maxCharCount = remainingBuffer / Utf8Helper.MaxBytePerChar;
            var remainingLength = cellValue.Length - cellValueIndex;
            var lastIteration = remainingLength <= maxCharCount;
            var length = lastIteration ? remainingLength : maxCharCount;
            Buffer.Advance(Utf8Helper.GetBytes(cellValue.AsSpan(cellValueIndex, length), Buffer.GetNextSpan()));
            cellValueIndex += length;
            return lastIteration;
        }
    }
}
