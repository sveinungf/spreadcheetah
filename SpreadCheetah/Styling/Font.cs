using SpreadCheetah.Helpers;
using System.Drawing;

namespace SpreadCheetah.Styling;

/// <summary>
/// Represents the font part of a <see cref="Style"/>.
/// </summary>
public sealed record Font
{
    internal const double DefaultSize = 11;

    private string? _name;

    /// <summary>Font name. Defaults to Calibri.</summary>
    public string? Name
    {
        get => _name;
        set => _name = value.WithEnsuredMaxLength(31);
    }

    /// <summary>Bold font weight. Defaults to <see langword="false"/>.</summary>
    public bool Bold { get; set; }

    /// <summary>Italic font type. Defaults to <see langword="false"/>.</summary>
    public bool Italic { get; set; }

    /// <summary>Adds a horizontal line through the center of the characters. Defaults to <see langword="false"/>.</summary>
    public bool Strikethrough { get; set; }

    /// <summary>Font size. Defaults to 11.</summary>
    public double Size { get; set; } = DefaultSize;

    /// <summary>ARGB (alpha, red, green, blue) color of the font.</summary>
    public Color? Color { get; set; }
}
