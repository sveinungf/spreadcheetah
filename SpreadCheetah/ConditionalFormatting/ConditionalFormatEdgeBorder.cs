using SpreadCheetah.Helpers;
using System.Drawing;

namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Represents an edge border part of a <see cref="ConditionalFormatStyle"/>.
/// </summary>
public sealed record ConditionalFormatEdgeBorder
{
    /// <summary>Border style. Defaults to None.</summary>
    public ConditionalFormatBorderStyle BorderStyle { get; set => field = Guard.DefinedEnumValue(value); }

    /// <summary>ARGB (alpha, red, green, blue) color of the border.</summary>
    public Color? Color { get; set; }

    internal bool IsDefault => this is
    {
        BorderStyle: ConditionalFormatBorderStyle.None,
        Color: null
    };
}
