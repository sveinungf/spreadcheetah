using SpreadCheetah.ConditionalFormatting.Internal;
using SpreadCheetah.Styling;

namespace SpreadCheetah.ConditionalFormatting;

public abstract class ConditionalFormatRule
{
    public static UniqueValuesFormatRuleBuilder UniqueValues() => new();

    internal ConditionalFormatStyle? Style { get; init; }

    internal abstract ImmutableConditionalFormatRule ToImmutable(int? styleDxfId);

    private protected ConditionalFormatRule()
    {
    }
}
