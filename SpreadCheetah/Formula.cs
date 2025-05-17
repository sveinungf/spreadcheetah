using SpreadCheetah.Formulas;

namespace SpreadCheetah;

/// <summary>
/// Represents a formula for a worksheet cell.
/// </summary>
public readonly record struct Formula
{
    internal string FormulaText { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="Formula"/> struct with a formula text.
    /// The formula text should be without a starting equal sign (=).
    /// </summary>
    public Formula(string? formulaText)
    {
        FormulaText = formulaText ?? "";
    }

    /// <summary>
    /// Creates a hyperlink formula that represents a link to the specified URI.
    /// The URI must be an absolute URI, and cannot exceed 255 characters in length.
    /// </summary>
    public static Formula Hyperlink(Uri uri) => HyperlinkFormula.From(uri);

    /// <summary>
    /// Creates a hyperlink formula that represents a link to the specified URI with a friendly display name.
    /// The URI must be an absolute URI. Both the URI and the friendly name cannot exceed 255 characters in length.
    /// </summary>
    public static Formula Hyperlink(Uri uri, string friendlyName) => HyperlinkFormula.From(uri, friendlyName);
}
