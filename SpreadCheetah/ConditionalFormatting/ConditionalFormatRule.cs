using SpreadCheetah.ConditionalFormatting.Internal;

namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Represents a conditional format rule for one or more worksheet cells.
/// </summary>
public abstract class ConditionalFormatRule
{
    /// <summary>
    /// Creates a conditional format rule that applies to cells with duplicate values.
    /// </summary>
    public static DuplicateValuesFormatRuleBuilder DuplicateValues() => new();

    /// <summary>
    /// Creates a conditional format rule that applies to cells with unique values.
    /// </summary>
    public static UniqueValuesFormatRuleBuilder UniqueValues() => new();

    /// <summary>
    /// Creates a conditional format rule that applies to cells where the specified formula evaluates to true.
    /// </summary>
    public static MatchesFormulaFormatRuleBuilder MatchesFormula(Formula formula) => new(formula);

    internal ConditionalFormatStyle? Style { get; init; }

    internal abstract InternalConditionalFormatRule ToInternal(int? styleDxfId);

    private protected ConditionalFormatRule()
    {
    }
}
