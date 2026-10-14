using SpreadCheetah.ConditionalFormatting.Internal;

namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Represents a conditional format rule that applies to cells based on comparing their numeric value to a range.
/// Use the builder from <see cref="ConditionalFormatRule.Values"/> to create an instance of this class.
/// </summary>
public sealed class NumericRangeFormatRule : ConditionalFormatRule
{
    internal double FirstValue { get; }
    internal double SecondValue { get; }
    internal bool IsBetween { get; }

    internal NumericRangeFormatRule(double firstValue, double secondValue, bool isBetween)
    {
        FirstValue = firstValue;
        SecondValue = secondValue;
        IsBetween = isBetween;
    }

    internal override InternalConditionalFormatRule ToInternal(int? styleDxfId)
    {
        return new InternalNumericRangeFormatRule
        {
            StyleDxfId = styleDxfId,
            FirstValue = FirstValue,
            SecondValue = SecondValue,
            IsBetween = IsBetween
        };
    }
}
