using SpreadCheetah.Helpers;

namespace SpreadCheetah.Styling
{
    public struct NumberFormat
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

        /// <inheritdoc />
        public override string ToString()
        {
            if (CustomFormat is not null) return CustomFormat;
            if (PredefinedFormat.HasValue) return Enum.GetName(typeof(PredefinedNumberFormat), PredefinedFormat.Value);
            return string.Empty;
        }

        /// <inheritdoc />
        public override bool Equals(object obj)
        {
            if (obj is NumberFormat numberFormat)
            {
                return string.Equals(numberFormat.CustomFormat, CustomFormat, StringComparison.Ordinal) &&
                       numberFormat.PredefinedFormat == PredefinedFormat;
            }
            return false;
        }

        /// <summary>
        /// Implcitly creates a predefined number format.
        /// </summary>
#pragma warning disable CA2225 // Operator overloads have named alternates - named alternatives are provided, named Custom and Predefined, for more fluent syntax.
        public static implicit operator NumberFormat(PredefinedNumberFormat format) => NumberFormat.Predefined(format);

        /// <summary>
        /// Implicitly creates a custom or predefined number format
        /// </summary>
        public static implicit operator NumberFormat(string formatString)
        {
            // Backwards compatibilty - if string being cast from is one of the old hard-coded predefined number format strings, convert it to an actual predefined number format to match previous behaviour
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
#pragma warning restore CA2225 // Operator overloads have named alternates
    }
}
