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

    public int ActualNumberOfColumns
    {
        get
        {
            if (Table.NumberOfColumns is { } numberOfColumns)
                return numberOfColumns;

            var headerLength = _headerNames.Length;
            if (headerLength == 0 && Table.TotalRowMaxColumnNumber is { } totalRowLength)
                return totalRowLength;

            return headerLength;
        }
    }

    public string GetHeaderName(int index)
    {
        Debug.Assert(index < ActualNumberOfColumns);
        var name = _headerNames.Length > 0
            ? _headerNames[index]
            : StringHelper.Invariant($"Column{index + 1}");

        Debug.Assert(!string.IsNullOrEmpty(name));
        return name;
    }

    public void SetHeaderNames(ReadOnlySpan<string> fullRowValues)
    {
        var tableOffset = FirstColumn - 1;

        if (fullRowValues.Length <= tableOffset)
            TableThrowHelper.MissingHeaderName();

        var headerNames = fullRowValues.Slice(tableOffset);
        Debug.Assert(!headerNames.IsEmpty);

        var expectedColumns = Table.NumberOfColumns ?? Table.TotalRowMaxColumnNumber;
        if (expectedColumns is { } columns && headerNames.Length < columns)
            TableThrowHelper.MissingHeaderNames(columns, headerNames.Length);

        if (Table.NumberOfColumns is { } numberOfColumns)
            headerNames = headerNames.Slice(0, numberOfColumns);

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
        var allColumnOptions = Table.ColumnOptions; // TODO: Can this be null?

        for (var i = 0; i < ActualNumberOfColumns; ++i)
        {
            var headerName = GetHeaderName(i);
            var columnOptions = allColumnOptions?.GetValueOrDefault(i + 1);

            var cell = columnOptions switch
            {
                { TotalRowLabel: { } label } => new Cell(label),
                { TotalRowFunction: { } func } => new Cell(TotalRowFormula(func, headerName)),
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