using System.Drawing;

namespace SpreadCheetah.Styling.Internal;

internal readonly record struct ImmutableFill
{
    public required Color? Color { get; init; }

    public bool IsDefault => Color is null;

    public static ImmutableFill From(Fill fill) => new() { Color = fill.Color };
}
