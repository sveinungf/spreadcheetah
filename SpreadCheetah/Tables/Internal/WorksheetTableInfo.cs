using SpreadCheetah.Helpers;

namespace SpreadCheetah.Tables.Internal;

internal sealed class WorksheetTableInfo
{
    private string?[]? _headerNames; // TODO: Array from array pool.

    public required ushort FirstColumn { get; init; }
    public required uint FirstRow { get; init; }
    public required Table Table { get; init; } // TODO: Make an immutable type (e.g. ImmutableTable like for styles)
    public uint? LastDataRow { get; set; }
    public bool Active => LastDataRow is null;

    public ReadOnlySpan<string?> HeaderNames => _headerNames.AsSpan();

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