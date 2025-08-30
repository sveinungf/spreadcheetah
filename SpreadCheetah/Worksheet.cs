using SpreadCheetah.CellReferences;
using SpreadCheetah.CellWriters;
using SpreadCheetah.Helpers;
using SpreadCheetah.Images.Internal;
using SpreadCheetah.MetadataXml;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using SpreadCheetah.Tables.Internal;
using SpreadCheetah.Validations;
using SpreadCheetah.Worksheets;
using SpreadCheetah.Worksheets.Internal;
using System.Runtime.CompilerServices;

namespace SpreadCheetah;

internal sealed class Worksheet : IDisposable, IAsyncDisposable
{
    private readonly Stream _stream;
#pragma warning disable CA2213 // Disposed by Spreadsheet
    private readonly SpreadsheetBuffer _buffer;
#pragma warning restore CA2213 // Disposed by Spreadsheet
    private readonly BaseRowWriter<Cell> _cellRowWriter;
    private readonly BaseRowWriter<DataCell> _dataCellRowWriter;
    private readonly BaseRowWriter<StyledCell> _styledCellRowWriter;
    private readonly CellWriterState _state;
    private readonly string? _autoFilterRange;
    private Dictionary<SingleCellOrCellRangeReference, DataValidation>? _validations;
    private HashSet<CellRangeRelativeReference>? _cellMerges;

    public Dictionary<SingleCellRelativeReference, string>? Notes { get; private set; }
    public List<WorksheetDimensionRun>? ColumnWidthRuns { get; }
    public List<WorksheetDimensionRun>? RowHeightRuns { get; private set; }
    public List<WorksheetTableInfo>? Tables { get; private set; }
    public List<WorksheetImage>? Images { get; private set; }
    public double? DefaultColumnWidth { get; }

    public Worksheet(Stream stream, DefaultStyling? defaultStyling, SpreadsheetBuffer buffer,
        bool writeCellReferenceAttributes, WorksheetOptions? options)
    {
        _stream = stream;
        _buffer = buffer;
        _state = new CellWriterState(buffer, defaultStyling);
        _autoFilterRange = options?.AutoFilter?.CellRange.Reference;
        ColumnWidthRuns = options?.GetColumnWidthRuns();
        DefaultColumnWidth = options?.DefaultColumnWidth;

        var columnStyles = options?.GetColumnStyles();
        if (writeCellReferenceAttributes)
        {
            _cellRowWriter = CreateRowWriter(CellWithReferenceWriter.Instance, columnStyles, _state);
            _dataCellRowWriter = CreateRowWriter(DataCellWithReferenceWriter.Instance, columnStyles, _state);
            _styledCellRowWriter = CreateRowWriter(StyledCellWithReferenceWriter.Instance, columnStyles, _state);
        }
        else
        {
            _cellRowWriter = CreateRowWriter(CellWriter.Instance, columnStyles, _state);
            _dataCellRowWriter = CreateRowWriter(DataCellWriter.Instance, columnStyles, _state);
            _styledCellRowWriter = CreateRowWriter(StyledCellWriter.Instance, columnStyles, _state);
        }
    }

    private static BaseRowWriter<T> CreateRowWriter<T>(
        ICellWriter<T> cellWriter,
        IReadOnlyDictionary<int, StyleId>? columnStyles,
        CellWriterState state)
        where T : struct
    {
        return columnStyles is null
            ? new RowWriter<T>(cellWriter, state)
            : new RowWithColumnStylesWriter<T>(cellWriter, state, columnStyles);
    }

    public int NextRowNumber => (int)_state.NextRowIndex;

    public void StartTable(ImmutableTable table, int firstColumnNumber)
    {
        if (_autoFilterRange is not null)
            TableThrowHelper.TablesNotAllowedOnWorksheetWithAutoFilter();

        var tables = Tables ??= [];

        if (tables.GetActive() is not null)
            TableThrowHelper.OnlyOneActiveTableAllowed();

        var worksheetTable = new WorksheetTableInfo
        {
            FirstColumn = (ushort)firstColumnNumber,
            FirstRow = _state.NextRowIndex,
            Table = table
        };

        tables.Add(worksheetTable);
    }

    public WorksheetTableInfo? GetActiveTable() => Tables?.GetActive();

    [OverloadResolutionPriority(1)]
    public bool TryAddRow(ReadOnlySpan<DataCell> cells, RowOptions? options) => TryAddRow(cells, _dataCellRowWriter, options);
    public bool TryAddRow(ReadOnlySpan<StyledCell> cells, RowOptions? options) => TryAddRow(cells, _styledCellRowWriter, options);
    public bool TryAddRow(ReadOnlySpan<Cell> cells, RowOptions? options) => TryAddRow(cells, _cellRowWriter, options);
    public bool TryAddRow(IList<DataCell> cells, RowOptions? options) => TryAddRow(cells, _dataCellRowWriter, options);
    public bool TryAddRow(IList<StyledCell> cells, RowOptions? options) => TryAddRow(cells, _styledCellRowWriter, options);
    public bool TryAddRow(IList<Cell> cells, RowOptions? options) => TryAddRow(cells, _cellRowWriter, options);

    private bool TryAddRow<TCell, TWriter>(ReadOnlySpan<TCell> cells, TWriter writer, RowOptions? options)
        where TCell : struct
        where TWriter : BaseRowWriter<TCell>
    {
        return options is null
            ? writer.TryAddRow(cells, _state.NextRowIndex++)
            : TryAddRowWithOptions(cells, _state.NextRowIndex++, writer, options);
    }

    private bool TryAddRow<TCell, TWriter>(IList<TCell> cells, TWriter writer, RowOptions? options)
        where TCell : struct
        where TWriter : BaseRowWriter<TCell>
    {
        return options is null
            ? writer.TryAddRow(cells, _state.NextRowIndex++)
            : TryAddRowWithOptions(cells, _state.NextRowIndex++, writer, options);
    }

    private bool TryAddRowWithOptions<TCell, TWriter>(ReadOnlySpan<TCell> cells, uint rowIndex, TWriter writer, RowOptions options)
        where TCell : struct
        where TWriter : BaseRowWriter<TCell>
    {
        var result = writer.TryAddRow(cells, rowIndex, options);
        if (result && options.Height is { } height)
            UpdateRowHeightRuns(rowIndex, height);

        return result;
    }

    private bool TryAddRowWithOptions<TCell, TWriter>(IList<TCell> cells, uint rowIndex, TWriter writer, RowOptions options)
        where TCell : struct
        where TWriter : BaseRowWriter<TCell>
    {
        var result = writer.TryAddRow(cells, rowIndex, options);
        if (result && options.Height is { } height)
            UpdateRowHeightRuns(rowIndex, height);

        return result;
    }

    private void UpdateRowHeightRuns(uint rowIndex, double height)
    {
        RowHeightRuns ??= [];

        if (RowHeightRuns is [.., var lastRun] && lastRun.TryContinueWith(rowIndex, height))
            return;

        var newRun = new WorksheetDimensionRun(rowIndex, height);
        RowHeightRuns.Add(newRun);
    }

    [OverloadResolutionPriority(1)]
    public ValueTask AddRowAsync(IList<DataCell> cells, RowOptions? options, CancellationToken ct) => AddRowAsync(cells, _dataCellRowWriter, options, ct);
    public ValueTask AddRowAsync(IList<StyledCell> cells, RowOptions? options, CancellationToken ct) => AddRowAsync(cells, _styledCellRowWriter, options, ct);
    public ValueTask AddRowAsync(IList<Cell> cells, RowOptions? options, CancellationToken ct) => AddRowAsync(cells, _cellRowWriter, options, ct);
    public ValueTask AddRowAsync(ReadOnlyMemory<DataCell> cells, RowOptions? options, CancellationToken ct) => AddRowAsync(cells, _dataCellRowWriter, options, ct);
    public ValueTask AddRowAsync(ReadOnlyMemory<StyledCell> cells, RowOptions? options, CancellationToken ct) => AddRowAsync(cells, _styledCellRowWriter, options, ct);
    public ValueTask AddRowAsync(ReadOnlyMemory<Cell> cells, RowOptions? options, CancellationToken ct) => AddRowAsync(cells, _cellRowWriter, options, ct);

    private ValueTask AddRowAsync<TCell, TWriter>(IList<TCell> cells, TWriter writer, RowOptions? options, CancellationToken ct)
        where TCell : struct
        where TWriter : BaseRowWriter<TCell>
    {
        return writer.AddRowAsync(cells, _state.NextRowIndex - 1, options, _stream, ct);
    }

    private ValueTask AddRowAsync<TCell, TWriter>(ReadOnlyMemory<TCell> cells, TWriter writer, RowOptions? options, CancellationToken ct)
        where TCell : struct
        where TWriter : BaseRowWriter<TCell>
    {
        return writer.AddRowAsync(cells, _state.NextRowIndex - 1, options, _stream, ct);
    }

    public bool TryAddDataValidation(string reference, DataValidation validation)
    {
        _validations ??= [];

        if (_validations.Count >= SpreadsheetConstants.MaxNumberOfDataValidations)
            return false;

        var cellReference = SingleCellOrCellRangeReference.Create(reference);
        _validations[cellReference] = validation;
        return true;
    }

    public void AddImage(in WorksheetImage image, out bool firstImage)
    {
        firstImage = false;

        if (Images is null)
        {
            firstImage = true;
            Images = [];
        }

        Images.Add(image);
    }

    public void AddNote(string cellReference, string noteText, out bool firstNote)
    {
        var reference = SingleCellRelativeReference.Create(cellReference);
        firstNote = false;

        if (Notes is null)
        {
            firstNote = true;
            Notes = [];
        }

        Notes[reference] = noteText;
    }

    public void MergeCells(CellRangeRelativeReference cellRange)
    {
        _cellMerges ??= [];
        _cellMerges.Add(cellRange);
    }

    public async ValueTask FinishTableAsync(bool throwWhenNoTable, CancellationToken token)
    {
        var tableInfo = Tables.GetActive();
        if (tableInfo is null && !throwWhenNoTable)
            return;
        if (tableInfo is null)
            TableThrowHelper.NoActiveTables();

        var tableRows = _state.NextRowIndex - tableInfo.FirstRow;
        if (tableRows == 0)
            TableThrowHelper.NoRows(tableInfo.Table.Name);
        if (tableInfo.ActualNumberOfColumns == 0)
            TableThrowHelper.NoColumns(tableInfo.Table.Name);

        var tableHasOnlyHeaderRow = tableRows == 1 && tableInfo.HasHeaderRow;
        if (tableHasOnlyHeaderRow && !TryAddRow([], null))
        {
            // Add an empty row so that the table doesn't cause an error when opening the file in Excel
            await AddRowAsync([], null, token).ConfigureAwait(false);
        }

        tableInfo.LastDataRow = _state.NextRowIndex - 1;

        if (!tableInfo.Table.HasTotalRow)
            return;

        var totalRow = tableInfo.CreateTotalRow();
        if (!TryAddRow(totalRow, null))
            await AddRowAsync(totalRow, null, token).ConfigureAwait(false);
    }

    public async ValueTask FinishAsync(CancellationToken token)
    {
        await FinishTableAsync(throwWhenNoTable: false, token).ConfigureAwait(false);

        using var cellMergesPooledArray = _cellMerges?.ToPooledArray();
        using var validationsPooledArray = _validations?.ToPooledArray();

        await WorksheetEndXml.WriteAsync(
            cellMerges: cellMergesPooledArray?.Memory,
            validations: validationsPooledArray?.Memory,
            autoFilterRange: _autoFilterRange,
            hasNotes: Notes is not null,
            hasImages: Images is not null,
            tableCount: Tables?.Count ?? 0,
            buffer: _buffer,
            stream: _stream,
            token: token).ConfigureAwait(false);

        await _stream.FlushAsync(token).ConfigureAwait(false);
    }

    public void Dispose() => _stream.Dispose();
    public ValueTask DisposeAsync() => _stream.DisposeAsync();
}
