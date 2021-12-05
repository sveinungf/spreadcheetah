using System.Drawing;
using System.Net;

namespace SpreadCheetah.Styling;

/// <summary>
/// Represents the font part of a <see cref="Style"/>.
/// </summary>
public sealed class Font : IEquatable<Font>
{
    private string? _name;

    /// <summary>Font name. Defaults to Calibri.</summary>
    public string? Name
    {
        get => _name;
        set => _name = WebUtility.HtmlEncode(value);
    }

    /// <summary>Bold font weight. Defaults to <c>false</c>.</summary>
    public bool Bold { get; set; }

    /// <summary>Italic font type. Defaults to <c>false</c>.</summary>
    public bool Italic { get; set; }

    /// <summary>Adds a horizontal line through the center of the characters. Defaults to <c>false</c>.</summary>
    public bool Strikethrough { get; set; }

    /// <summary>Font size. Defaults to 11.</summary>
    public double Size { get; set; } = 11;

    /// <summary>ARGB (alpha, red, green, blue) color of the font.</summary>
    public Color? Color { get; set; }

    /// <inheritdoc/>
    public bool Equals(Font? other) => other != null
        && string.Equals(Name, other.Name, StringComparison.Ordinal)
        && Bold == other.Bold && Italic == other.Italic && Strikethrough == other.Strikethrough
        && Size == other.Size
        && EqualityComparer<Color?>.Default.Equals(Color, other.Color);

    /// <inheritdoc/>
    public override bool Equals(object? obj) => obj is Font other && Equals(other);

    /// <inheritdoc/>
    public override int GetHashCode() => HashCode.Combine(Name, Bold, Italic, Strikethrough, Size, Color);
}
