namespace SpreadCheetah.ConditionalFormatting;

public abstract class ConditionalFormatRule
{
    private protected ConditionalFormatRule()
    {
    }

    public static UniqueValuesFormatRuleBuilder UniqueValues() => new();
}
