using SpreadCheetah.Helpers;

namespace SpreadCheetah.Tables.Internal;

internal sealed class WorksheetTableInfo
{
    private string?[]? _headerNames;

    /// <summary>The first column of the table. 1-based.</summary>
    public required ushort FirstColumn { get; init; }

    /// <summary>The first row of the table. 1-based.</summary>
    public required uint FirstRow { get; init; }
    public required ImmutableTable Table { get; init; }
    public uint? LastDataRow { get; set; }
    public bool Active => LastDataRow is null;
    public bool HasHeaderRow => !HeaderNames.IsEmpty; // TODO: Not entirely correct. If calling "AddHeaderRow" with an empty row, then this should return true.

    public ReadOnlySpan<string?> HeaderNames => _headerNames.AsSpan();
    public int ActualNumberOfColumns => Table.NumberOfColumns ?? HeaderNames.Length;

    public void SetHeaderNames(ReadOnlySpan<string?> fullRowValues)
    {
        var startIndex = FirstColumn - 1;
        if (startIndex >= fullRowValues.Length)
            return;

        var startingValues = fullRowValues.Slice(startIndex);
        var startingValuesLength = fullRowValues.Length - startIndex;

        _headerNames = Table.NumberOfColumns is { } count && count < startingValuesLength
            ? startingValues.Slice(0, count).ToArray()
            : startingValues.ToArray();
    }

    public void SetHeaderNames(IList<string?> fullRowValues)
    {
        var startIndex = FirstColumn - 1;
        if (startIndex >= fullRowValues.Count)
            return;

        var startingValues = fullRowValues.Skip(startIndex);
        var startingValuesLength = fullRowValues.Count - startIndex;

        _headerNames = Table.NumberOfColumns is { } count && count < startingValuesLength
            ? startingValues.Take(count).ToArray()
            : startingValues.ToArray();
    }

    public List<Cell> CreateTotalRow()
    {
        var cells = new List<Cell>();
        var allColumnOptions = Table.ColumnOptions; // TODO: What to do when this is null?
        var headers = HeaderNames;

        for (var i = 0; i < headers.Length; ++i)
        {
            var columnOptions = allColumnOptions?.GetValueOrDefault(i + 1) ?? new();

            var cell = columnOptions switch
            {
                { TotalRowLabel: { } label } => new Cell(label),
                { TotalRowFunction: { } func } => new Cell(TotalRowFormula(func, headers[i])),
                _ => new Cell()
            };

            cells.Add(cell);
        }

        return cells;
    }

    private Formula TotalRowFormula(TableTotalRowFunction function, string? headerName)
    {
        var functionNumber = function switch
        {
            TableTotalRowFunction.Average => 101,
            TableTotalRowFunction.Count => 103,
            TableTotalRowFunction.CountNumbers => 102,
            TableTotalRowFunction.Maximum => 104,
            TableTotalRowFunction.Minimum => 105,
            TableTotalRowFunction.StandardDeviation => 107,
            TableTotalRowFunction.Sum => 109,
            TableTotalRowFunction.Variance => 110,
            _ => 0
        };

        if (functionNumber == 0)
            return new Formula(); // TODO: Expected behavior?

        // TODO: What if headerName is null?
        // TODO: Do we need to adjust the tableName in any way?
        // TODO: Do we need to adjust the headerName in any way?
        var formulaText = StringHelper.Invariant($"SUBTOTAL({functionNumber},{Table.Name}[{headerName}])");
        return new Formula(formulaText);
    }
}