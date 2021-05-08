using System;
using System.Net;

namespace SpreadCheetah
{
    /// <summary>
    /// Represents a formula for a worksheet cell.
    /// </summary>
    public readonly struct Formula : IEquatable<Formula>
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

        /// <inheritdoc/>
        public bool Equals(Formula other)
        {
            return string.Equals(FormulaText, other.FormulaText, StringComparison.Ordinal);
        }

        /// <inheritdoc/>
        public override bool Equals(object? obj) => obj is Formula formula && Equals(formula);

        /// <inheritdoc/>
        public override int GetHashCode() => HashCode.Combine(FormulaText);

        /// <summary>Determines whether two instances have the same value.</summary>
        public static bool operator ==(Formula left, Formula right) => left.Equals(right);

        /// <summary>Determines whether two instances have different values.</summary>
        public static bool operator !=(Formula left, Formula right) => !left.Equals(right);
    }
}
