using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Worksheets;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadCheetah
{
    internal sealed class Worksheet : IDisposable, IAsyncDisposable
    {
        private const string SheetHeader =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">";

        private const string ColsBegin = "<cols>";
        private const string ColsEnd = "</cols>";
        private const string SheetDataBegin = "<sheetData>";
        private const string ColsEndSheetDataBegin = ColsEnd + SheetDataBegin;
        private const string SheetFooter = "</sheetData></worksheet>";

        private readonly Stream _stream;
        private readonly SpreadsheetBuffer _buffer;
        private readonly DataCellWriter _dataCellWriter;
        private readonly StyledCellWriter _styledCellWriter;
        private int _nextRowIndex;

        public Worksheet(Stream stream, SpreadsheetBuffer buffer)
        {
            _stream = stream;
            _buffer = buffer;
            _dataCellWriter = new DataCellWriter(buffer);
            _styledCellWriter = new StyledCellWriter(buffer);
            _nextRowIndex = 1;
        }

        public async ValueTask WriteHeadAsync(WorksheetOptions? options, CancellationToken token)
        {
            _buffer.Index += Utf8Helper.GetBytes(SheetHeader, _buffer.GetNextSpan());
            if (options == null)
            {
                _buffer.Index += Utf8Helper.GetBytes(SheetDataBegin, _buffer.GetNextSpan());
                return;
            }

            _buffer.Index += Utf8Helper.GetBytes(ColsBegin, _buffer.GetNextSpan());

            var sb = new StringBuilder();
            foreach (var keyValuePair in options.ColumnOptions)
            {
                var column = keyValuePair.Value;
                if (column.Width == null) continue;

                sb.Clear();
                sb.Append("<col min=\"");
                sb.Append(keyValuePair.Key);
                sb.Append("\" max=\"");
                sb.Append(keyValuePair.Key);
                sb.Append("\" width=\"");
                sb.Append(column.Width.Value);
                sb.Append("\" customWidth=\"1\" />");

                if (sb.Length > _buffer.GetRemainingBuffer())
                    await _buffer.FlushToStreamAsync(_stream, token).ConfigureAwait(false);

                _buffer.Index += Utf8Helper.GetBytes(sb.ToString(), _buffer.GetNextSpan());
            }

            if (ColsEndSheetDataBegin.Length > _buffer.GetRemainingBuffer())
                await _buffer.FlushToStreamAsync(_stream, token).ConfigureAwait(false);

            _buffer.Index += Utf8Helper.GetBytes(ColsEndSheetDataBegin, _buffer.GetNextSpan());
        }

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
