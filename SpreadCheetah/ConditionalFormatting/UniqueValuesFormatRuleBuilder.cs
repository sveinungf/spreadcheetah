namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Builder for a unique values conditional format rule.
/// Use <see cref="ConditionalFormatRule.UniqueValues"/> to create an instance of this builder.
/// </summary>
public sealed class UniqueValuesFormatRuleBuilder
{
    internal UniqueValuesFormatRuleBuilder()
    {
    }

    /// <summary>
    /// Sets the style that should be applied to cells with unique values.
    /// </summary>
    public UniqueValuesFormatRule WithStyle(ConditionalFormatStyle style)
    {
        ArgumentNullException.ThrowIfNull(style);
        return new UniqueValuesFormatRule { Style = style };
    }
}
