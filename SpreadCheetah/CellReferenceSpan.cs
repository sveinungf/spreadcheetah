namespace SpreadCheetah;

[Flags]
internal enum CellReferenceSpan
{
    None = 0,
    SingleCell = 1,
    CellRange = 2,
    SingleCellOrCellRange = SingleCell | CellRange
}
