using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.MetadataXml;
using SpreadCheetah.Styling.Internal;
using SpreadCheetah.Validations;
using SpreadCheetah.Worksheets;

namespace SpreadCheetah;

internal sealed class Worksheet : IDisposable, IAsyncDisposable
{
    private readonly Stream _stream;
    private readonly SpreadsheetBuffer _buffer;
    private readonly CellWriter _cellWriter;
    private readonly DataCellWriter _dataCellWriter;
    private readonly StyledCellWriter _styledCellWriter;
    private readonly CellWriterState _state;
    private Dictionary<CellReference, DataValidation>? _validations;
    private HashSet<CellReference>? _cellMerges;
    private string? _autoFilterRange;

    public Worksheet(Stream stream, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer)
    {
        _stream = stream;
        _buffer = buffer;
        _state = new CellWriterState(buffer, false);
        _cellWriter = new CellWriter(_state, defaultStyling);
        _dataCellWriter = new DataCellWriter(_state, defaultStyling);
        _styledCellWriter = new StyledCellWriter(_state, defaultStyling);
    }

    public int NextRowNumber => (int)_state.NextRowIndex;

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

        if (options?.AutoFilter is not null)
        {
            _autoFilterRange = options.AutoFilter.CellRange.Reference;
        }
    }

    public bool TryAddRow(IList<Cell> cells)
        => _cellWriter.TryAddRow(cells, _state.NextRowIndex++);
    public bool TryAddRow(IList<DataCell> cells)
        => _dataCellWriter.TryAddRow(cells, _state.NextRowIndex++);
    public bool TryAddRow(IList<StyledCell> cells)
        => _styledCellWriter.TryAddRow(cells, _state.NextRowIndex++);
    public bool TryAddRow(IList<Cell> cells, RowOptions options)
        => _cellWriter.TryAddRow(cells, _state.NextRowIndex++, options);
    public bool TryAddRow(IList<DataCell> cells, RowOptions options)
        => _dataCellWriter.TryAddRow(cells, _state.NextRowIndex++, options);
    public bool TryAddRow(IList<StyledCell> cells, RowOptions options)
        => _styledCellWriter.TryAddRow(cells, _state.NextRowIndex++, options);
    public bool TryAddRow(ReadOnlySpan<Cell> cells)
        => _cellWriter.TryAddRow(cells, _state.NextRowIndex++);
    public bool TryAddRow(ReadOnlySpan<DataCell> cells)
        => _dataCellWriter.TryAddRow(cells, _state.NextRowIndex++);
    public bool TryAddRow(ReadOnlySpan<StyledCell> cells)
        => _styledCellWriter.TryAddRow(cells, _state.NextRowIndex++);
    public bool TryAddRow(ReadOnlySpan<Cell> cells, RowOptions options)
        => _cellWriter.TryAddRow(cells, _state.NextRowIndex++, options);
    public bool TryAddRow(ReadOnlySpan<DataCell> cells, RowOptions options)
        => _dataCellWriter.TryAddRow(cells, _state.NextRowIndex++, options);
    public bool TryAddRow(ReadOnlySpan<StyledCell> cells, RowOptions options)
        => _styledCellWriter.TryAddRow(cells, _state.NextRowIndex++, options);
    public ValueTask AddRowAsync(IList<Cell> cells, CancellationToken ct)
        => _cellWriter.AddRowAsync(cells, _state.NextRowIndex - 1, _stream, ct);
    public ValueTask AddRowAsync(IList<DataCell> cells, CancellationToken ct)
        => _dataCellWriter.AddRowAsync(cells, _state.NextRowIndex - 1, _stream, ct);
    public ValueTask AddRowAsync(IList<StyledCell> cells, CancellationToken ct)
        => _styledCellWriter.AddRowAsync(cells, _state.NextRowIndex - 1, _stream, ct);
    public ValueTask AddRowAsync(IList<Cell> cells, RowOptions options, CancellationToken ct)
        => _cellWriter.AddRowAsync(cells, _state.NextRowIndex - 1, options, _stream, ct);
    public ValueTask AddRowAsync(IList<DataCell> cells, RowOptions options, CancellationToken ct)
        => _dataCellWriter.AddRowAsync(cells, _state.NextRowIndex - 1, options, _stream, ct);
    public ValueTask AddRowAsync(IList<StyledCell> cells, RowOptions options, CancellationToken ct)
        => _styledCellWriter.AddRowAsync(cells, _state.NextRowIndex - 1, options, _stream, ct);
    public ValueTask AddRowAsync(ReadOnlyMemory<Cell> cells, CancellationToken ct)
        => _cellWriter.AddRowAsync(cells, _state.NextRowIndex - 1, _stream, ct);
    public ValueTask AddRowAsync(ReadOnlyMemory<DataCell> cells, CancellationToken ct)
        => _dataCellWriter.AddRowAsync(cells, _state.NextRowIndex - 1, _stream, ct);
    public ValueTask AddRowAsync(ReadOnlyMemory<StyledCell> cells, CancellationToken ct)
        => _styledCellWriter.AddRowAsync(cells, _state.NextRowIndex - 1, _stream, ct);
    public ValueTask AddRowAsync(ReadOnlyMemory<Cell> cells, RowOptions options, CancellationToken ct)
        => _cellWriter.AddRowAsync(cells, _state.NextRowIndex - 1, options, _stream, ct);
    public ValueTask AddRowAsync(ReadOnlyMemory<DataCell> cells, RowOptions options, CancellationToken ct)
        => _dataCellWriter.AddRowAsync(cells, _state.NextRowIndex - 1, options, _stream, ct);
    public ValueTask AddRowAsync(ReadOnlyMemory<StyledCell> cells, RowOptions options, CancellationToken ct)
        => _styledCellWriter.AddRowAsync(cells, _state.NextRowIndex - 1, options, _stream, ct);

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
