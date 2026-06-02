namespace SpreadCheetah.ConditionalFormatting;

public sealed class UniqueValuesFormatRuleBuilder
{
    internal UniqueValuesFormatRuleBuilder()
    {
    }

    public UniqueValuesFormatRule WithStyle(ConditionalFormatStyle style)
    {
        ArgumentNullException.ThrowIfNull(style);
        return new UniqueValuesFormatRule { Style = style };
    }
}
