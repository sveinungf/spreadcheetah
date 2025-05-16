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

    public static Formula Hyperlink(Uri uri) => HyperlinkFormula.From(uri);

    public static Formula Hyperlink(Uri uri, string friendlyName) => HyperlinkFormula.From(uri, friendlyName);
}
