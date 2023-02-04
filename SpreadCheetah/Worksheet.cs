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
    private const string SheetHeader =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">";

    private const string SheetDataBegin = "<sheetData>";

    private readonly Stream _stream;
    private readonly SpreadsheetBuffer _buffer;
    private readonly CellWriter _cellWriter;
    private readonly DataCellWriter _dataCellWriter;
    private readonly StyledCellWriter _styledCellWriter;
    private uint _nextRowIndex;
    private Dictionary<CellReference, DataValidation>? _validations;
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
        _buffer.Advance(Utf8Helper.GetBytes(SheetHeader, _buffer.GetSpan()));
        if (options is null)
        {
            _buffer.Advance(Utf8Helper.GetBytes(SheetDataBegin, _buffer.GetSpan()));
            return;
        }

        var sb = new StringBuilder();

        if (options.FrozenColumns is not null || options.FrozenRows is not null)
        {
            await _buffer.FlushToStreamAsync(_stream, token).ConfigureAwait(false);
            WriteSheetViewsXml(sb, options);
            await _buffer.WriteAsciiStringAsync(sb.ToString(), _stream, token).ConfigureAwait(false);
            sb.Clear();
        }

        await WriteColsXmlAsync(sb, options, token).ConfigureAwait(false);

        _buffer.Advance(Utf8Helper.GetBytes(SheetDataBegin, _buffer.GetSpan()));

        if (options.AutoFilter is not null)
        {
            _autoFilterRange = options.AutoFilter.Range;
        }
    }

    private static void WriteSheetViewsXml(StringBuilder sb, WorksheetOptions options)
    {
        sb.Append("<sheetViews><sheetView workbookViewId=\"0\"><pane ");

        if (options.FrozenColumns is not null)
            sb.Append("xSplit=\"").Append(options.FrozenColumns.Value).Append("\" ");

        if (options.FrozenRows is not null)
            sb.Append("ySplit=\"").Append(options.FrozenRows.Value).Append("\" ");

        sb.Append("topLeftCell=\"");
        sb.AppendColumnName((options.FrozenColumns ?? 0) + 1);
        sb.Append((options.FrozenRows ?? 0) + 1);
        sb.Append("\" activePane=\"");

        if (options.FrozenColumns is not null && options.FrozenRows is not null)
            sb.Append("bottomRight");
        else if (options.FrozenColumns is not null)
            sb.Append("topRight");
        else
            sb.Append("bottomLeft");

        sb.Append("\" state=\"frozen\"/>");

        if (options.FrozenColumns is not null && options.FrozenRows is not null)
            sb.Append("<selection pane=\"bottomRight\"/>");
        if (options.FrozenRows is not null)
            sb.Append("<selection pane=\"bottomLeft\"/>");
        if (options.FrozenColumns is not null)
            sb.Append("<selection pane=\"topRight\"/>");

        sb.Append("</sheetView></sheetViews>");
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

    public void AddDataValidation(CellReference reference, DataValidation validation)
    {
        _validations ??= new Dictionary<CellReference, DataValidation>();
        _validations[reference] = validation;
    }

    public async ValueTask FinishAsync(CancellationToken token)
    {
        await _buffer.WriteAsciiStringAsync("</sheetData>", _stream, token).ConfigureAwait(false);

        if (_validations is not null)
            await DataValidationXml.WriteAsync(_stream, _buffer, _validations, token).ConfigureAwait(false);

        if (_autoFilterRange is not null)
            await AutoFilterXml.WriteAsync(_stream, _buffer, _autoFilterRange, token).ConfigureAwait(false);

        await _buffer.WriteAsciiStringAsync("</worksheet>", _stream, token).ConfigureAwait(false);
        await _buffer.FlushToStreamAsync(_stream, token).ConfigureAwait(false);
        await _stream.FlushAsync(token).ConfigureAwait(false);
    }

    public void Dispose() => _stream.Dispose();
    public ValueTask DisposeAsync() => _stream.DisposeAsync();
}
