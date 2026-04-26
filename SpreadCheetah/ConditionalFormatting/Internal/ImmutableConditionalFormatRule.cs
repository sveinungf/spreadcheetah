namespace SpreadCheetah.ConditionalFormatting.Internal;

internal abstract record ImmutableConditionalFormatRule
{
    public int? StyleDxfId { get; init; }

    public abstract bool TryWrite(SpreadsheetBuffer buffer, int priority);
}
