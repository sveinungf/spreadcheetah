using System.Net;

namespace SpreadCheetah
{
    public readonly struct Formula
    {
        // TODO: Without '='
        public string FormulaText { get; }

        public Formula(string? formulaText)
        {
            FormulaText = formulaText != null ? WebUtility.HtmlEncode(formulaText) : string.Empty;
        }
    }
}
