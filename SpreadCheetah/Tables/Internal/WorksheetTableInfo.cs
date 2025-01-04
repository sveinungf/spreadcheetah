using SpreadCheetah.Helpers;

namespace SpreadCheetah.Tables.Internal;

internal sealed class WorksheetTableInfo
{
    private string?[]? _headerNames;

    /// <summary>The first column of the table. 1-based.</summary>
    public required ushort FirstColumn { get; init; }

    /// <summary>The first row of the table. 1-based.</summary>
    public required uint FirstRow { get; init; }
    public required Table Table { get; init; } // TODO: Make an immutable type (e.g. ImmutableTable like for styles)
    public uint? LastDataRow { get; set; }
    public bool Active => LastDataRow is null;

    public ReadOnlySpan<string?> HeaderNames => _headerNames.AsSpan();

    public void SetHeaderNames(ReadOnlySpan<string?> fullRowValues)
    {
        if (!fullRowValues.TrySlice(FirstColumn - 1, out var startingValues))
            return;
        if (startingValues.IsEmpty)
            return;

        // TODO: Consider using an array from the array pool.
        _headerNames = Table.NumberOfColumns is { } count && count < startingValues.Length
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

        // TODO: Consider using an array from the array pool.
        _headerNames = Table.NumberOfColumns is { } count && count < startingValuesLength
            ? startingValues.Take(count).ToArray()
            : startingValues.ToArray();
    }

    public PooledArray<Cell> CreateTotalRow()
    {
        // TODO: Example
        //var cells = new List<Cell>
        //{
        //    new Cell("Total"),
        //    new Cell(),
        //    new Cell(new Formula("SUBTOTAL(109,Table1[[Count ]])")),
        //    new Cell(new Formula("SUBTOTAL(101,Table1[Price])")),
        //};
        throw new NotImplementedException();
    }
}