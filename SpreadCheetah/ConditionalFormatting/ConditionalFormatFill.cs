using System.Drawing;

namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Represents the fill part of a <see cref="ConditionalFormatStyle"/>.
/// </summary>
public sealed record ConditionalFormatFill
{
    /// <summary>ARGB (alpha, red, green, blue) color of the fill.</summary>
    public Color? Color { get; set; }

    internal bool IsDefault => Color is null;
}
