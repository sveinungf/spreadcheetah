using SpreadCheetah.CellReferences;
using SpreadCheetah.MetadataXml.Attributes;

namespace SpreadCheetah.ConditionalFormatting.Internal;

internal sealed record InternalNumericComparisonFormatRule : InternalConditionalFormatRule
{
    public required double Value { get; init; }
    public required NumericComparisonOperator Operator { get; init; }

    public override bool TryWrite(SpreadsheetBuffer buffer, int priority, SimpleSingleCellReference topLeftCell)
    {
        var dxfIdAttribute = new IntAttribute("dxfId"u8, StyleDxfId);
        var priorityAttribute = new IntAttribute("priority"u8, priority);
        var operatorAttribute = new SpanByteAttribute("operator"u8, GetOperatorValue(Operator));

        return buffer.TryWrite(
            $"{"<cfRule type=\"cellIs\""u8}" +
            $"{dxfIdAttribute}" +
            $"{priorityAttribute}" +
            $"{operatorAttribute}" +
            $"{"><formula>"u8}" +
            $"{Value}" +
            $"{"</formula></cfRule>"u8}");
    }

    private static ReadOnlySpan<byte> GetOperatorValue(NumericComparisonOperator op) => op switch
    {
        NumericComparisonOperator.GreaterThan => "greaterThan"u8,
        NumericComparisonOperator.LessThan => "lessThan"u8,
        NumericComparisonOperator.EqualTo => "equal"u8,
        NumericComparisonOperator.NotEqualTo => "notEqual"u8,
        NumericComparisonOperator.GreaterThanOrEqualTo => "greaterThanOrEqual"u8,
        _ => "lessThanOrEqual"u8
    };
}
