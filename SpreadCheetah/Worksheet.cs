using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.MetadataXml;
using SpreadCheetah.Styling.Internal;
using SpreadCheetah.Validations;
using SpreadCheetah.Worksheets;
using System.Text;

namespace SpreadCheetah;

internal sealed class Worksheet : IDisposable, IAsyncDisposable
{
    private const string SheetDataBegin = "<sheetData>";

    private readonly Stream _stream;
    private readonly SpreadsheetBuffer _buffer;
    private readonly CellWriter _cellWriter;
    private readonly DataCellWriter _dataCellWriter;
    private readonly StyledCellWriter _styledCellWriter;
    private uint _nextRowIndex;
    private Dictionary<CellReference, DataValidation>? _validations;
    private HashSet<CellReference>? _cellMerges;
    private string? _autoFilterRange;

    public Worksheet(Stream stream, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        _stream = stream;
        _buffer = buffer;
        _cellWriter = new CellWriter(buffer, defaultStyling);
        _dataCellWriter = new DataCellWriter(buffer, defaultStyling);
        _styledCellWriter = new StyledCellWriter(buffer, defaultStyling);
        _nextRowIndex = 1;
    }

    public int NextRowNumber => (int)_nextRowIndex;

    public async ValueTask WriteHeadAsync(WorksheetOptions? options, CancellationToken token)
    {
        var writer = new WorksheetStartXml(options);
        bool done;
        var buffer = _buffer;
        var stream = _stream;

        do
        {
            done = writer.TryWrite(buffer.GetSpan(), out var bytesWritten);
            buffer.Advance(bytesWritten);
            await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        } while (!done);

        if (options is null)
            return;

        var sb = new StringBuilder();
        await WriteColsXmlAsync(sb, options, token).ConfigureAwait(false);

        _buffer.Advance(Utf8Helper.GetBytes(SheetDataBegin, _buffer.GetSpan()));

        if (options.AutoFilter is not null)
        {
            _autoFilterRange = options.AutoFilter.CellRange.Reference;
        }
    }

    private async ValueTask WriteColsXmlAsync(StringBuilder sb, WorksheetOptions options, CancellationToken token)
    {
        var firstColumnWritten = false;
        foreach (var keyValuePair in options.ColumnOptions)
        {
            var column = keyValuePair.Value;
            if (column.Width is null)
                continue;

            if (!firstColumnWritten)
            {
                sb.Append("<cols>");
                firstColumnWritten = true;
            }

            sb.Append("<col min=\"");
            sb.Append(keyValuePair.Key);
            sb.Append("\" max=\"");
            sb.Append(keyValuePair.Key);
            sb.Append("\" width=\"");
            sb.AppendDouble(column.Width.Value);
            sb.Append("\" customWidth=\"1\" />");

            await _buffer.WriteAsciiStringAsync(sb.ToString(), _stream, token).ConfigureAwait(false);
            sb.Clear();
        }

        // The <cols> tag should only be written if there is a column with a real width set.
        if (firstColumnWritten)
            _buffer.Advance(Utf8Helper.GetBytes("</cols>", _buffer.GetSpan()));
    }

    public bool TryAddRow(IList<Cell> cells, out int currentIndex)
        => _cellWriter.TryAddRow(cells, _nextRowIndex++, out currentIndex);
    public bool TryAddRow(IList<DataCell> cells, out int currentIndex)
        => _dataCellWriter.TryAddRow(cells, _nextRowIndex++, out currentIndex);
    public bool TryAddRow(IList<StyledCell> cells, out int currentIndex)
        => _styledCellWriter.TryAddRow(cells, _nextRowIndex++, out currentIndex);
    public bool TryAddRow(IList<Cell> cells, RowOptions options, out int currentIndex)
        => _cellWriter.TryAddRow(cells, _nextRowIndex++, options, out currentIndex);
    public bool TryAddRow(IList<DataCell> cells, RowOptions options, out int currentIndex)
        => _dataCellWriter.TryAddRow(cells, _nextRowIndex++, options, out currentIndex);
    public bool TryAddRow(IList<StyledCell> cells, RowOptions options, out int currentIndex)
        => _styledCellWriter.TryAddRow(cells, _nextRowIndex++, options, out currentIndex);
    public bool TryAddRow(ReadOnlySpan<Cell> cells, out int currentIndex)
        => _cellWriter.TryAddRow(cells, _nextRowIndex++, out currentIndex);
    public bool TryAddRow(ReadOnlySpan<DataCell> cells, out int currentIndex)
        => _dataCellWriter.TryAddRow(cells, _nextRowIndex++, out currentIndex);
    public bool TryAddRow(ReadOnlySpan<StyledCell> cells, out int currentIndex)
        => _styledCellWriter.TryAddRow(cells, _nextRowIndex++, out currentIndex);
    public bool TryAddRow(ReadOnlySpan<Cell> cells, RowOptions options, out int currentIndex)
        => _cellWriter.TryAddRow(cells, _nextRowIndex++, options, out currentIndex);
    public bool TryAddRow(ReadOnlySpan<DataCell> cells, RowOptions options, out int currentIndex)
        => _dataCellWriter.TryAddRow(cells, _nextRowIndex++, options, out currentIndex);
    public bool TryAddRow(ReadOnlySpan<StyledCell> cells, RowOptions options, out int currentIndex)
        => _styledCellWriter.TryAddRow(cells, _nextRowIndex++, options, out currentIndex);
    public ValueTask AddRowAsync(IList<Cell> cells, int currentIndex, CancellationToken ct)
        => _cellWriter.AddRowAsync(cells, _nextRowIndex - 1, currentIndex, _stream, ct);
    public ValueTask AddRowAsync(IList<DataCell> cells, int currentIndex, CancellationToken ct)
        => _dataCellWriter.AddRowAsync(cells, _nextRowIndex - 1, currentIndex, _stream, ct);
    public ValueTask AddRowAsync(IList<StyledCell> cells, int currentIndex, CancellationToken ct)
        => _styledCellWriter.AddRowAsync(cells, _nextRowIndex - 1, currentIndex, _stream, ct);
    public ValueTask AddRowAsync(IList<Cell> cells, RowOptions options, int currentIndex, CancellationToken ct)
        => _cellWriter.AddRowAsync(cells, _nextRowIndex - 1, options, currentIndex, _stream, ct);
    public ValueTask AddRowAsync(IList<DataCell> cells, RowOptions options, int currentIndex, CancellationToken ct)
        => _dataCellWriter.AddRowAsync(cells, _nextRowIndex - 1, options, currentIndex, _stream, ct);
    public ValueTask AddRowAsync(IList<StyledCell> cells, RowOptions options, int currentIndex, CancellationToken ct)
        => _styledCellWriter.AddRowAsync(cells, _nextRowIndex - 1, options, currentIndex, _stream, ct);
    public ValueTask AddRowAsync(ReadOnlyMemory<Cell> cells, int currentIndex, CancellationToken ct)
        => _cellWriter.AddRowAsync(cells, _nextRowIndex - 1, currentIndex, _stream, ct);
    public ValueTask AddRowAsync(ReadOnlyMemory<DataCell> cells, int currentIndex, CancellationToken ct)
        => _dataCellWriter.AddRowAsync(cells, _nextRowIndex - 1, currentIndex, _stream, ct);
    public ValueTask AddRowAsync(ReadOnlyMemory<StyledCell> cells, int currentIndex, CancellationToken ct)
        => _styledCellWriter.AddRowAsync(cells, _nextRowIndex - 1, currentIndex, _stream, ct);
    public ValueTask AddRowAsync(ReadOnlyMemory<Cell> cells, RowOptions options, int currentIndex, CancellationToken ct)
        => _cellWriter.AddRowAsync(cells, _nextRowIndex - 1, options, currentIndex, _stream, ct);
    public ValueTask AddRowAsync(ReadOnlyMemory<DataCell> cells, RowOptions options, int currentIndex, CancellationToken ct)
        => _dataCellWriter.AddRowAsync(cells, _nextRowIndex - 1, options, currentIndex, _stream, ct);
    public ValueTask AddRowAsync(ReadOnlyMemory<StyledCell> cells, RowOptions options, int currentIndex, CancellationToken ct)
        => _styledCellWriter.AddRowAsync(cells, _nextRowIndex - 1, options, currentIndex, _stream, ct);

    public bool TryAddDataValidation(string reference, DataValidation validation)
    {
        _validations ??= new Dictionary<CellReference, DataValidation>();

        if (_validations.Count >= SpreadsheetConstants.MaxNumberOfDataValidations)
            return false;

        var cellReference = CellReference.Create(reference, true, CellReferenceType.RelativeOrAbsolute);
        _validations[cellReference] = validation;
        return true;
    }

    public void MergeCells(CellReference cellRange)
    {
        _cellMerges ??= new HashSet<CellReference>();
        _cellMerges.Add(cellRange);
    }

    public async ValueTask FinishAsync(CancellationToken token)
    {
        var writer = new WorksheetEndXml(_cellMerges?.ToList(), _validations?.ToList(), _autoFilterRange);
        bool done;
        var buffer = _buffer;
        var stream = _stream;

        do
        {
            done = writer.TryWrite(buffer.GetSpan(), out var bytesWritten);
            buffer.Advance(bytesWritten);
            await buffer.FlushToStreamAsync(stream, token).ConfigureAwait(false);
        } while (!done);

        await _stream.FlushAsync(token).ConfigureAwait(false);
    }

    public void Dispose() => _stream.Dispose();
    public ValueTask DisposeAsync() => _stream.DisposeAsync();
}
