namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Values that specifies the style of borders used in conditional formatting rules.
/// </summary>
public enum ConditionalFormatBorderStyle
{
    /// <summary>No border.</summary>
    None,

    /// <summary>Thin solid line.</summary>
    Thin = 1,

    /// <summary>Dashed line.</summary>
    Dashed = 3,

    /// <summary>Dotted line.</summary>
    Dotted = 4,

    /// <summary>Hairline.</summary>
    Hair = 7,

    /// <summary>Dash-dot line.</summary>
    DashDot = 9,

    /// <summary>Dash-dot-dot line.</summary>
    DashDotDot = 11
}
