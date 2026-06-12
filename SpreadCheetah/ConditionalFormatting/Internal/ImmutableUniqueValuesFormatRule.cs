using SpreadCheetah.MetadataXml.Attributes;

namespace SpreadCheetah.ConditionalFormatting.Internal;

internal sealed record ImmutableUniqueValuesFormatRule : ImmutableConditionalFormatRule
{
    public override bool TryWrite(SpreadsheetBuffer buffer, int priority)
    {
        var dxfIdAttribute = new IntAttribute("dxfId"u8, StyleDxfId);
        var priorityAttribute = new IntAttribute("priority"u8, priority);

        return buffer.TryWrite(
            $"{"<cfRule type=\"uniqueValues\""u8}" +
            $"{dxfIdAttribute}" +
            $"{priorityAttribute}" +
            $"{"/>"u8}");
    }
}
