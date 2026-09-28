using SpreadCheetah.ConditionalFormatting.Internal;

namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Represents a conditional format rule that applies to cells where the specified formula evaluates to true.
/// Use the builder from <see cref="ConditionalFormatRule.MatchesFormula"/> to create an instance of this class.
/// </summary>
public sealed class MatchesFormulaFormatRule : ConditionalFormatRule
{
    internal Formula Formula { get; }

    internal MatchesFormulaFormatRule(Formula formula)
    {
        Formula = formula;
    }

    internal override InternalConditionalFormatRule ToInternal(int? styleDxfId)
    {
        return new InternalMatchesFormulaFormatRule
        {
            StyleDxfId = styleDxfId,
            Formula = Formula
        };
    }
}
