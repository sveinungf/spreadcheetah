using System.Drawing;

namespace SpreadCheetah.Styling.Internal;

internal readonly record struct ImmutableFill(Color? Color)
{
    public static ImmutableFill From(Fill fill) => new(fill.Color);
}
