using SpreadCheetah.CellReferences;

namespace SpreadCheetah.ConditionalFormatting.Internal;

internal abstract record ImmutableConditionalFormatRule
{
    // TODO: Verify if both relative and absolute references are supported in conditional formatting rules.
    public required SingleCellOrCellRangeReference Reference { get; init; }
}
