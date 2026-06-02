namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Represents all border parts of a <see cref="ConditionalFormatStyle"/>.
/// </summary>
public sealed record ConditionalFormatBorder
{
    // TODO: Validate that these are not set to null in setters.
    /// <summary>Left border for the cell.</summary>
    public ConditionalFormatEdgeBorder Left { get; set; } = new();

    /// <summary>Right border for the cell.</summary>
    public ConditionalFormatEdgeBorder Right { get; set; } = new();

    /// <summary>Top border for the cell.</summary>
    public ConditionalFormatEdgeBorder Top { get; set; } = new();

    /// <summary>Bottom border for the cell.</summary>
    public ConditionalFormatEdgeBorder Bottom { get; set; } = new();

    internal bool IsDefault => this is
    {
        Left.IsDefault: true,
        Right.IsDefault: true,
        Top.IsDefault: true,
        Bottom.IsDefault: true
    };
}
