namespace SpreadCheetah.Styling;

/// <summary>
/// Represents style for one or more worksheet cells.
/// </summary>
public sealed record Style
{
    /// <summary>Alignment for the cell.</summary>
    public Alignment Alignment { get; set; } = new();

    /// <summary>Border for the cell.</summary>
    public Border Border { get; set; } = new();

    /// <summary>Fill for the cell.</summary>
    public Fill Fill { get; set; } = new();

    /// <summary>Font for the cell's value.</summary>
    public Font Font { get; set; } = new();

    /// <summary>Format that defines how a number or <see cref="DateTime"/> cell should be displayed.</summary>
    [Obsolete($"Use {nameof(Style)}.{nameof(Format)} instead")]
    public string? NumberFormat
    {
        get => Format?.CustomFormat;
        set => Format = value == null ? null : Styling.NumberFormat.FromLegacyString(value);
    }

    /// <summary>Format that defines how a number or <see cref="DateTime"/> cell should be displayed.</summary>
    public NumberFormat? Format { get; set; }
}
