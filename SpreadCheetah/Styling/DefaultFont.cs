using SpreadCheetah.Helpers;

namespace SpreadCheetah.Styling;

/// <summary>
/// Represents the default font for worksheets.
/// </summary>
public sealed record DefaultFont
{
    internal const string DefaultName = "Calibri";
    internal const double DefaultSize = 11;

    /// <summary>Font name. Defaults to Calibri.</summary>
    public string? Name { get; set => field = Guard.FontNameLengthInRange(value); }

    /// <summary>Font size. Defaults to 11.</summary>
    public double Size { get; set => field = Guard.FontSizeInRange(value); } = DefaultSize;
}
