using SpreadCheetah.Styling;

namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Values that specifies the style of borders used in conditional formatting rules.
/// </summary>
public enum ConditionalFormatBorderStyle
{
    /// <summary>No border.</summary>
    None = BorderStyle.None,

    /// <summary>Thin solid line.</summary>
    Thin = BorderStyle.Thin,

    /// <summary>Dashed line.</summary>
    Dashed = BorderStyle.Dashed,

    /// <summary>Dotted line.</summary>
    Dotted = BorderStyle.Dotted,

    /// <summary>Hairline.</summary>
    Hair = BorderStyle.Hair,

    /// <summary>Dash-dot line.</summary>
    DashDot = BorderStyle.DashDot,

    /// <summary>Dash-dot-dot line.</summary>
    DashDotDot = BorderStyle.DashDotDot
}
