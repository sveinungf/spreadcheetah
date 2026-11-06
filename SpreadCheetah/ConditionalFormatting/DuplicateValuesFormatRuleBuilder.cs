namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Builder for a duplicate values conditional format rule.
/// Use <see cref="ConditionalFormatRule.DuplicateValues"/> to create an instance of this builder.
/// </summary>
public sealed class DuplicateValuesFormatRuleBuilder
{
    internal DuplicateValuesFormatRuleBuilder()
    {
    }

    /// <summary>
    /// Sets the style that should be applied to cells with duplicate values.
    /// </summary>
    /// <returns>The conditional format rule.</returns>
    public DuplicateValuesFormatRule WithStyle(ConditionalFormatStyle style)
    {
        ArgumentNullException.ThrowIfNull(style);
        return new DuplicateValuesFormatRule { Style = style };
    }
}
