using System.Net;

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
        FormulaText = formulaText != null ? WebUtility.HtmlEncode(formulaText) : string.Empty;
    }
}
