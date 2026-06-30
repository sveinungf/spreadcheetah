using SpreadCheetah.CellWriters;
using SpreadCheetah.Formulas;

namespace SpreadCheetah;

/// <summary>
/// Represents a formula for a worksheet cell.
/// </summary>
public readonly record struct Formula
{
    internal string FormulaText { get; }

    internal bool IsR1C1 { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="Formula"/> struct with a formula text.
    /// The formula text should be without a starting equal sign (=).
    /// </summary>
    public Formula(string? formulaText)
    {
        FormulaText = formulaText ?? "";
    }

    private Formula(string formulaText, bool isR1C1)
    {
        FormulaText = formulaText;
        IsR1C1 = isR1C1;
    }

    /// <summary>
    /// Creates a formula from text in the R1C1 reference style. References in the formula will be converted to the
    /// A1 reference style (which is required in the XLSX file format) based on the position of the cell that the
    /// formula is added to. This makes it possible to reuse the same formula for multiple cells.
    /// The formula text should be without a starting equal sign (=).
    /// </summary>
    public static Formula R1C1(string formula)
    {
        ArgumentNullException.ThrowIfNull(formula);
        return new Formula(formula, isR1C1: true);
    }

    internal string GetFormulaText(CellWriterState state) => IsR1C1
        ? R1C1FormulaConverter.ToA1(FormulaText, (int)(state.NextRowIndex - 1), state.Column + 1)
        : FormulaText;

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
