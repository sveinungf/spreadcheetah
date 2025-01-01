namespace SpreadCheetah.Tables.Internal;

internal sealed class WorksheetTableInfo
{
    private string?[]? _headerNames; // TODO: Array from array pool.

    public required ushort FirstColumn { get; init; }
    public required uint FirstRow { get; init; }
    public uint? LastDataRow { get; set; }
    public bool Active => LastDataRow is null;

    public ReadOnlySpan<string?> HeaderNames => _headerNames.AsSpan();
}