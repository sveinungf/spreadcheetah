using SpreadCheetah.Styling;

namespace SpreadCheetah.ConditionalFormatting.Internal;

internal sealed record ImmutableUniqueValuesFormatRule : ImmutableConditionalFormatRule
{
    public required Style Style { get; init; }
}
