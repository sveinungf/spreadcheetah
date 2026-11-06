using SpreadCheetah.Styling;

namespace SpreadCheetah.ConditionalFormatting;

/// <summary>
/// Values that specifies the type of underline used for conditional format fonts.
/// </summary>
public enum ConditionalFormatUnderline
{
    /// <summary>No underline.</summary>
    None = Underline.None,

    /// <summary>Single-line underlining.</summary>
    Single = Underline.Single,

    /// <summary>Double-line underlining.</summary>
    Double = Underline.Double
}
