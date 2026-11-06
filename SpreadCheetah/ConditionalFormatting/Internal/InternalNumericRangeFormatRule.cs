using SpreadCheetah.CellReferences;
using SpreadCheetah.MetadataXml.Attributes;

namespace SpreadCheetah.ConditionalFormatting.Internal;

internal sealed class InternalNumericRangeFormatRule : InternalConditionalFormatRule
{
    public required double FirstValue { get; init; }
    public required double SecondValue { get; init; }
    public required bool IsBetween { get; init; }

    public override bool TryWrite(SpreadsheetBuffer buffer, int priority, SimpleSingleCellReference topLeftCell)
    {
        var dxfIdAttribute = new IntAttribute("dxfId"u8, StyleDxfId);
        var priorityAttribute = new IntAttribute("priority"u8, priority);
        var operatorAttribute = new SpanByteAttribute("operator"u8, IsBetween ? "between"u8 : "notBetween"u8);

        return buffer.TryWrite(
            $"{"<cfRule type=\"cellIs\""u8}" +
            $"{dxfIdAttribute}" +
            $"{priorityAttribute}" +
            $"{operatorAttribute}" +
            $"{"><formula>"u8}" +
            $"{FirstValue}" +
            $"{"</formula><formula>"u8}" +
            $"{SecondValue}" +
            $"{"</formula></cfRule>"u8}");
    }
}
