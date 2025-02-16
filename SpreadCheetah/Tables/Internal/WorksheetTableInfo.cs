using SpreadCheetah.Helpers;
using System.Diagnostics;

namespace SpreadCheetah.Tables.Internal;

internal sealed class WorksheetTableInfo
{
    private string[] _headerNames = [];

    /// <summary>The first column of the table. 1-based.</summary>
    public required ushort FirstColumn { get; init; }

    /// <summary>The first row of the table. 1-based.</summary>
    public required uint FirstRow { get; init; }
    public required ImmutableTable Table { get; init; }
    public uint? LastDataRow { get; set; }
    public bool Active => LastDataRow is null;
    public bool HasHeaderRow => _headerNames.Length > 0;

    public ReadOnlySpan<string> HeaderNames => _headerNames.AsSpan();
    public int ActualNumberOfColumns => Table.NumberOfColumns ?? HeaderNames.Length;

    public void SetHeaderNames(ReadOnlySpan<string> fullRowValues)
    {
        var tableOffset = FirstColumn - 1;

        if (fullRowValues.Length <= tableOffset)
            TableThrowHelper.MissingHeaderName();

        var headerNames = fullRowValues.Slice(tableOffset);
        Debug.Assert(!headerNames.IsEmpty);

        if (Table.NumberOfColumns is { } numberOfColumns)
        {
            if (headerNames.Length < numberOfColumns)
                TableThrowHelper.MissingHeaderNames(numberOfColumns, headerNames.Length);

            headerNames = headerNames.Slice(0, numberOfColumns);
        }

        var uniqueHeaderNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < headerNames.Length; i++)
        {
            var headerName = headerNames[i];
            if (string.IsNullOrEmpty(headerName))
                TableThrowHelper.HeaderNameNullOrEmpty(i + 1);
            if (char.IsWhiteSpace(headerName[0]) || char.IsWhiteSpace(headerName[^1]))
                TableThrowHelper.HeaderNameCanNotBeginOrEndWithWhitespace();
            if (!uniqueHeaderNames.Add(headerName))
                TableThrowHelper.DuplicateHeaderName(headerName);
        }

        _headerNames = headerNames.ToArray();
    }

    public List<Cell> CreateTotalRow()
    {
        var cells = new List<Cell>();
        var allColumnOptions = Table.ColumnOptions; // TODO: What to do when this is null?
        var headers = HeaderNames;

        for (var i = 0; i < headers.Length; ++i)
        {
            Debug.Assert(!string.IsNullOrEmpty(headers[i]));
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

    private Formula TotalRowFormula(TableTotalRowFunction function, string headerName)
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

        // TODO: Do we need to adjust the tableName in any way?
        var formulaText = StringHelper.Invariant($"SUBTOTAL({functionNumber},{Table.Name}[{headerName}])");
        return new Formula(formulaText);
    }
}