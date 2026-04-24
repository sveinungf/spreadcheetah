using SpreadCheetah.Styling;

namespace SpreadCheetah.ConditionalFormatting;

public sealed class UniqueValuesFormatRule : ConditionalFormatRule
{
    internal Style Style { get; init; }

    internal UniqueValuesFormatRule(Style style)
    {
        Style = style;
    }
}
