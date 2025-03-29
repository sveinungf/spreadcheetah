namespace SpreadCheetah.Styling;

/// <summary>
/// Values that specifies the type of underline used for fonts.
/// </summary>
public enum Underline
{
    /// <summary>No underline.</summary>
    None,

    /// <summary>Single-line underlining.</summary>
    Single,

    /// <summary>Single-line accounting underlining. The underline is drawn further down, under the descenders of characters.</summary>
    SingleAccounting,

    /// <summary>Double-line underlining.</summary>
    Double,

    /// <summary>Double-line accounting underlining. The underline is drawn further down, under the descenders of characters.</summary>
    DoubleAccounting
}
