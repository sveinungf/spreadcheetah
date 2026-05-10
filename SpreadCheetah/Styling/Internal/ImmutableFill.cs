using SpreadCheetah.ConditionalFormatting;
using System.Drawing;

namespace SpreadCheetah.Styling.Internal;

internal readonly record struct ImmutableFill(Color? Color)
{
    public static ImmutableFill From(Fill fill) => new(fill.Color);

    public static ImmutableFill From(ConditionalFormatFill fill) => new(fill.Color);
}
