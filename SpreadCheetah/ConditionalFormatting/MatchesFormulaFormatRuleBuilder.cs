using SpreadCheetah.Helpers;

namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Builder for a formula conditional format rule.
/// Use <see cref="ConditionalFormatRule.MatchesFormula"/> to create an instance of this builder.
/// </summary>
public sealed class MatchesFormulaFormatRuleBuilder
{
    internal Formula Formula { get; }

    internal MatchesFormulaFormatRuleBuilder(Formula formula)
    {
        if (formula.FormulaText.Length == 0)
            ThrowHelper.FormulaEmpty(nameof(formula));

        Formula = formula;
    }

    /// <summary>
    /// Sets the style that should be applied to cells that match the formula.
    /// </summary>
    /// <param name="style">The conditional format style.</param>
    /// <returns>The conditional format rule.</returns>
    public MatchesFormulaFormatRule WithStyle(ConditionalFormatStyle style)
    {
        ArgumentNullException.ThrowIfNull(style);
        return new MatchesFormulaFormatRule(Formula) { Style = style };
    }
}
