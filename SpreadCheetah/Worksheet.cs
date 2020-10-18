using SpreadCheetah.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadCheetah
{
    internal sealed class Worksheet : IDisposable, IAsyncDisposable
    {
        private const string SheetHeader =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<sheetData>";

        private const string SheetFooter = "</sheetData></worksheet>";

        private const int Utf8MaxBytePerChar = 6;

        private static readonly int RowStartMaxByteCount = RowStart.Length + SpreadsheetConstants.RowIndexMaxDigits + RowStartEndTag.Length;

        private static ReadOnlySpan<byte> RowStart => new[]
        {
            (byte)'<', (byte)'r', (byte)'o', (byte)'w', (byte)' ', (byte)'r', (byte)'=', (byte)'"'
        };

        private static ReadOnlySpan<byte> RowStartEndTag => new[]
        {
            (byte)'"', (byte)'>'
        };

        private static ReadOnlySpan<byte> RowEnd => new[]
        {
            (byte)'<', (byte)'/', (byte)'r', (byte)'o', (byte)'w', (byte)'>'
        };

        private readonly Stream _stream;
        private readonly SpreadsheetBuffer _buffer;
        private int _nextRowIndex;

        private Worksheet(Stream stream, SpreadsheetBuffer buffer)
        {
            _stream = stream;
            _nextRowIndex = 1;
            _buffer = buffer;
        }

        public static Worksheet Create(Stream stream, SpreadsheetBuffer buffer)
        {
            var worksheet = new Worksheet(stream, buffer);
            worksheet.WriteHead();
            return worksheet;
        }

        private void WriteHead() => _buffer.Index += Utf8Helper.GetBytes(SheetHeader, _buffer.GetNextSpan());

        private static int GetRowStartBytes(int rowIndex, Span<byte> bytes)
        {
            var bytesWritten = SpanHelper.GetBytes(RowStart, bytes);
            bytesWritten += Utf8Helper.GetBytes(rowIndex, bytes.Slice(bytesWritten));
            bytesWritten += SpanHelper.GetBytes(RowStartEndTag, bytes.Slice(bytesWritten));
            return bytesWritten;
        }

        private bool TryWriteCell(in Cell cell, out int bytesNeeded)
        {
            bytesNeeded = 0;
            var remainingBuffer = _buffer.GetRemainingBuffer();

            // Try with an approximate cell value length
            var cellValueLength = cell.Value.Length * Utf8MaxBytePerChar;
            if (CellSpanHelper.MaxCellElementLength + cellValueLength < remainingBuffer)
            {
                _buffer.Index += CellSpanHelper.GetBytes(cell, _buffer.GetNextSpan());
                return true;
            }

            // Try with a more accurate cell value length
            cellValueLength = Utf8Helper.GetByteCount(cell.Value);
            bytesNeeded = CellSpanHelper.MaxCellElementLength + cellValueLength;
            if (bytesNeeded < remainingBuffer)
            {
                _buffer.Index += CellSpanHelper.GetBytes(cell, _buffer.GetNextSpan());
                return true;
            }

            return false;
        }

        private bool FinishWritingCellValue(string cellValue, ref int cellValueIndex)
        {
            var remainingBuffer = _buffer.GetRemainingBuffer();
            var maxCharCount = remainingBuffer / Utf8MaxBytePerChar;
            var remainingLength = cellValue.Length - cellValueIndex;
            var lastIteration = remainingLength <= maxCharCount;
            var length = lastIteration ? remainingLength : maxCharCount;
            _buffer.Index += Utf8Helper.GetBytes(cellValue.AsSpan(cellValueIndex, length), _buffer.GetNextSpan());
            cellValueIndex += length;
            return lastIteration;
        }

        public bool TryAddRow(IList<Cell> cells, out int currentIndex)
        {
            // Assuming previous actions on the worksheet ensured space in the buffer for row start
            _buffer.Index += GetRowStartBytes(_nextRowIndex++, _buffer.GetNextSpan());

            for (currentIndex = 0; currentIndex < cells.Count; ++currentIndex)
            {
                // Write cell if it fits in the buffer
                if (!TryWriteCell(cells[currentIndex], out _))
                    return false;
            }

            // Also ensuring space in the buffer for the next row start, so that we don't need to check space in the buffer twice
            if (RowEnd.Length + RowStartMaxByteCount > _buffer.GetRemainingBuffer())
                return false;

            _buffer.Index += SpanHelper.GetBytes(RowEnd, _buffer.GetNextSpan());
            return true;
        }

        public async ValueTask AddRowAsync(IList<Cell> cells, int currentIndex, CancellationToken token)
        {
            // If we get here that means that the next cell didn't fit in the buffer, so just flush right away
            await _buffer.FlushToStreamAsync(_stream, token).ConfigureAwait(false);

            for (var i = currentIndex; i < cells.Count; ++i)
            {
                var cell = cells[i];

                // Write cell if it fits in the buffer
                if (TryWriteCell(cell, out var bytesNeeded))
                    continue;

                await _buffer.FlushToStreamAsync(_stream, token).ConfigureAwait(false);

                // Write cell if it fits in the buffer
                if (bytesNeeded < _buffer.GetRemainingBuffer())
                {
                    _buffer.Index += CellSpanHelper.GetBytes(cell, _buffer.GetNextSpan());
                    continue;
                }

                // Write start element
                _buffer.Index += CellSpanHelper.GetStartElementBytes(cell.DataType, _buffer.GetNextSpan());

                // Write as much as possible from cell value
                var cellValueIndex = 0;
                while (!FinishWritingCellValue(cell.Value, ref cellValueIndex))
                {
                    await _buffer.FlushToStreamAsync(_stream, token).ConfigureAwait(false);
                }

                // Flush if can't fit the longest cell end element
                if (CellSpanHelper.MaxCellEndElementLength > _buffer.GetRemainingBuffer())
                    await _buffer.FlushToStreamAsync(_stream, token).ConfigureAwait(false);

                // Write end element directly
                _buffer.Index += CellSpanHelper.GetEndElementBytes(cell.DataType, _buffer.GetNextSpan());
            }

            // Also ensuring space in the buffer for the next row start, so that we don't need to check space in the buffer twice
            if (RowEnd.Length + RowStartMaxByteCount > _buffer.GetRemainingBuffer())
                await _buffer.FlushToStreamAsync(_stream, token).ConfigureAwait(false);

            _buffer.Index += SpanHelper.GetBytes(RowEnd, _buffer.GetNextSpan());
        }

        public async ValueTask FinishAsync(CancellationToken token)
        {
            if (Utf8Helper.GetByteCount(SheetFooter) > _buffer.GetRemainingBuffer())
                await _buffer.FlushToStreamAsync(_stream, token).ConfigureAwait(false);

            _buffer.Index += Utf8Helper.GetBytes(SheetFooter, _buffer.GetNextSpan());

            await _buffer.FlushToStreamAsync(_stream, token).ConfigureAwait(false);
            await _stream.FlushAsync(token).ConfigureAwait(false);
        }

        public ValueTask DisposeAsync()
        {
#if NETSTANDARD2_0
            Dispose();
            return default;
#else
            return _stream.DisposeAsync();
#endif
        }

        public void Dispose() => _stream.Dispose();
    }
}
