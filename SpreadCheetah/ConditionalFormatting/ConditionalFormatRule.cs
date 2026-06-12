using SpreadCheetah.ConditionalFormatting.Internal;

namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Represents a conditional format rule for one or more worksheet cells.
/// </summary>
public abstract class ConditionalFormatRule
{
    /// <summary>
    /// Creates a conditional format rule that applies to cells with unique values.
    /// </summary>
    public static UniqueValuesFormatRuleBuilder UniqueValues() => new();

    internal ConditionalFormatStyle? Style { get; init; }

    internal abstract ImmutableConditionalFormatRule ToImmutable(int? styleDxfId);

    private protected ConditionalFormatRule()
    {
    }
}
