using SpreadCheetah.Helpers;

namespace SpreadCheetah.Styling;

/// <summary>
/// Represents the alignment parts of a <see cref="Style"/>.
/// </summary>
public sealed record Alignment
{
    /// <summary>Allow text to be shown on multiple lines in the cell.</summary>
    public bool WrapText { get; set; }

    /// <summary>Horizontal alignment for the cell.</summary>
    public HorizontalAlignment Horizontal { get; set => field = Guard.DefinedEnumValue(value); }

    /// <summary>Vertical alignment for the cell.</summary>
    public VerticalAlignment Vertical { get; set => field = Guard.DefinedEnumValue(value); }

    /// <summary>
    /// Indentation for the cell. The value represents the number of character widths in Excel.
    /// The value can not be negative.
    /// </summary>
    public int Indent { get; set => field = Guard.NotNegative(value); }
}
