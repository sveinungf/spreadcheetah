using SpreadCheetah.CellReferences;
using SpreadCheetah.MetadataXml.Attributes;

namespace SpreadCheetah.ConditionalFormatting.Internal;

internal sealed class InternalAverageComparisonFormatRule : InternalConditionalFormatRule
{
    internal AverageComparisonOperator Operator { get; set; }

    public override bool TryWrite(SpreadsheetBuffer buffer, int priority, SimpleSingleCellReference topLeftCell)
    {
        var dxfIdAttribute = new IntAttribute("dxfId"u8, StyleDxfId);
        var priorityAttribute = new IntAttribute("priority"u8, priority);
        var aboveAverageAttribute = new BooleanAttribute("aboveAverage"u8, IsAboveOperator(Operator));
        var equalAverageAttribute = new BooleanAttribute("equalAverage"u8, IsEqualOperator(Operator));
        var standardDeviationsAttribute = new SpanByteAttribute("stdDev"u8, GetStandardDeviations(Operator));

        return buffer.TryWrite(
            $"{"<cfRule type=\"aboveAverage\""u8}" +
            $"{dxfIdAttribute}" +
            $"{priorityAttribute}" +
            $"{aboveAverageAttribute}" +
            $"{equalAverageAttribute}" +
            $"{standardDeviationsAttribute}" +
            $"{"/>"u8}");
    }

    private static bool? IsAboveOperator(AverageComparisonOperator op) => op switch
    {
        AverageComparisonOperator.Below => false,
        AverageComparisonOperator.EqualToOrBelow => false,
        AverageComparisonOperator.OneStdDevBelow => false,
        AverageComparisonOperator.TwoStdDevBelow => false,
        AverageComparisonOperator.ThreeStdDevBelow => false,
        _ => null
    };

    private static bool? IsEqualOperator(AverageComparisonOperator op) => op switch
    {
        AverageComparisonOperator.EqualToOrAbove => true,
        AverageComparisonOperator.EqualToOrBelow => true,
        _ => null
    };

    private static ReadOnlySpan<byte> GetStandardDeviations(AverageComparisonOperator op) => op switch
    {
        AverageComparisonOperator.OneStdDevAbove => "1"u8,
        AverageComparisonOperator.OneStdDevBelow => "1"u8,
        AverageComparisonOperator.TwoStdDevAbove => "2"u8,
        AverageComparisonOperator.TwoStdDevBelow => "2"u8,
        AverageComparisonOperator.ThreeStdDevAbove => "3"u8,
        AverageComparisonOperator.ThreeStdDevBelow => "3"u8,
        _ => []
    };
}
