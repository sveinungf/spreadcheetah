namespace SpreadCheetah.ConditionalFormatting.Internal;

internal sealed record ImmutableUniqueValuesFormatRule : ImmutableConditionalFormatRule
{
    public override bool TryWrite(SpreadsheetBuffer buffer, int priority)
    {
        if (StyleDxfId is not { } dxfId)
            throw new InvalidOperationException(); // TODO: Throw helper, or throw only for debug?

        return buffer.TryWrite(
            $"{"<cfRule type=\"uniqueValues\" dxfId=\""u8}" +
            $"{dxfId}" +
            $"{"\" priority=\""u8}" +
            $"{priority}" +
            $"{"\"/>"u8}");
    }
}
