using System.Drawing;

namespace SpreadCheetah.Styling.Internal;

internal readonly record struct ImmutableFill(Color? Color)
{
    public static ImmutableFill From(Fill fill) => new(fill.Color);

    public static ImmutableFill From(DifferentialFill fill) => new(fill.Color);
}
