using SpreadCheetah.CellReferences;
using SpreadCheetah.MetadataXml.Attributes;

namespace SpreadCheetah.ConditionalFormatting.Internal;

internal sealed record InternalUniqueValuesFormatRule : InternalConditionalFormatRule
{
    public override bool TryWrite(SpreadsheetBuffer buffer, int priority, SimpleSingleCellReference topLeftCell)
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
