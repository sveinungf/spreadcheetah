using SpreadCheetah.Helpers;
using System.Globalization;

namespace SpreadCheetah.Styling
{
    /// <summary>
    /// Format that defines how a number or <see cref="DateTime"/> cell should be displayed.
    /// May be a custom format, initialised by <see cref="NumberFormat.Custom"/>, or a predefined format, initialised by <see cref="NumberFormat.Predefined"/>
    /// </summary>
#pragma warning disable CA1815, CA2231 // Override equals and operator equals on value types - equals operator not appropriate for this type
    public readonly struct NumberFormat : IEquatable<NumberFormat>
#pragma warning restore CA1815, CA2231
    {
        private NumberFormat(string? customFormat, int? predefinedFormat)
        {
            CustomFormat = customFormat.WithEnsuredMaxLength(255);
            PredefinedFormat = predefinedFormat;
        }
        
        internal string? CustomFormat { get; }
        internal int? PredefinedFormat { get; }

        /// <summary>
        /// Creates a custom number format. The <paramref name="formatString"/> must be an <see href="https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68">Excel Format Code</see>
        /// </summary>
        /// <param name="formatString">The custom format string for this number format</param>
        /// <returns>A <see cref="NumberFormat"/> representing this custom format string</returns>
        public static NumberFormat Custom(string formatString)
        {
            return new NumberFormat(formatString, null);
        }

        /// <summary>
        /// Creates a predefined number format.
        /// </summary>
        /// <param name="format">The predefined format to use for this number format</param>
        /// <returns>A <see cref="NumberFormat"/> representing this predefined format</returns>
        public static NumberFormat Predefined(PredefinedNumberFormat format)
        {
            return new NumberFormat(null, (int)format);
        }

        /// <summary>
        /// Creates a number format from a string which may be custom or predefined.
        /// For backwards compatibility purposes only.
        /// </summary>
        /// <param name="formatString">The custom or predefined string to use for this number format</param>
        /// <returns>A <see cref="NumberFormat"/> representing this format</returns>
        internal static NumberFormat FromLegacyString(string formatString)
        {
            var predefinedNumberFormat = NumberFormats.GetPredefinedNumberFormatId(formatString);
            if (predefinedNumberFormat.HasValue)
            {
                return new NumberFormat(null, predefinedNumberFormat);
            }
            else
            {
                return new NumberFormat(formatString, null);
            }
        }

        /// <inheritdoc />
        public override string ToString()
        {
            if (CustomFormat is not null) return CustomFormat;
            if (PredefinedFormat.HasValue) return Enum.GetName(typeof(PredefinedNumberFormat), PredefinedFormat.Value) ?? PredefinedFormat.Value.ToString(CultureInfo.InvariantCulture);
            return string.Empty;
        }

        /// <inheritdoc />
        public override readonly bool Equals(object? obj)
        {
            if (obj is NumberFormat numberFormat)
            {
                return Equals(numberFormat);
            }
            return false;
        }

        /// <inheritdoc />
        public bool Equals(NumberFormat other)
        {
            return string.Equals(other.CustomFormat, CustomFormat, StringComparison.Ordinal) &&
                   other.PredefinedFormat == PredefinedFormat;
        }

        /// <inheritdoc />
        public override int GetHashCode()
        {
            return HashCode.Combine(CustomFormat, PredefinedFormat);
        }
    }
}
