using SpreadCheetah.Styling;

namespace SpreadCheetah.ConditionalFormatting;

public sealed class UniqueValuesFormatRuleBuilder
{
    internal UniqueValuesFormatRuleBuilder()
    {
    }

    public UniqueValuesFormatRule WithStyle(DifferentialStyle style) => new() { Style = style };
}
