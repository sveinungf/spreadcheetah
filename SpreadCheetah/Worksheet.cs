using SpreadCheetah.CellWriters;
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

        private readonly Stream _stream;
        private readonly SpreadsheetBuffer _buffer;
        private readonly DataCellWriter _dataCellWriter;
        private readonly StyledCellWriter _styledCellWriter;
        private int _nextRowIndex;

        private Worksheet(Stream stream, SpreadsheetBuffer buffer)
        {
            _stream = stream;
            _buffer = buffer;
            _dataCellWriter = new DataCellWriter(buffer);
            _styledCellWriter = new StyledCellWriter(buffer);
            _nextRowIndex = 1;
        }

        public static Worksheet Create(Stream stream, SpreadsheetBuffer buffer)
        {
            var worksheet = new Worksheet(stream, buffer);
            worksheet.WriteHead();
            return worksheet;
        }

        private void WriteHead() => _buffer.Index += Utf8Helper.GetBytes(SheetHeader, _buffer.GetNextSpan());
        public bool TryAddRow(IList<Cell> cells, out int currentIndex) => _dataCellWriter.TryAddRow(cells, _nextRowIndex++, out currentIndex);
        public bool TryAddRow(IList<StyledCell> cells, out int currentIndex) => _styledCellWriter.TryAddRow(cells, _nextRowIndex++, out currentIndex);
        public ValueTask AddRowAsync(IList<Cell> cells, int startIndex, CancellationToken ct) => _dataCellWriter.AddRowAsync(cells, startIndex, _stream, ct);
        public ValueTask AddRowAsync(IList<StyledCell> cells, int startIndex, CancellationToken ct) => _styledCellWriter.AddRowAsync(cells, startIndex, _stream, ct);

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
