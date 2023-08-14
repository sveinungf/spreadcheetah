using SpreadCheetah.Styling;
using Xunit;

namespace SpreadCheetah.Test.Tests
{
    public static class NumberFormatTests
    {
        [Theory]
        [InlineData(NumberFormats.General, StandardNumberFormat.General)]
        [InlineData(NumberFormats.Fraction, StandardNumberFormat.Fraction)]
        [InlineData(NumberFormats.FractionTwoDenominatorPlaces, StandardNumberFormat.FractionTwoDenominatorPlaces)]
        [InlineData(NumberFormats.NoDecimalPlaces, StandardNumberFormat.NoDecimalPlaces)]
        [InlineData(NumberFormats.Percent, StandardNumberFormat.Percent)]
        [InlineData(NumberFormats.PercentTwoDecimalPlaces, StandardNumberFormat.PercentTwoDecimalPlaces)]
        [InlineData(NumberFormats.Scientific, StandardNumberFormat.Scientific)]
        [InlineData(NumberFormats.ThousandsSeparator, StandardNumberFormat.ThousandsSeparator)]
        [InlineData(NumberFormats.ThousandsSeparatorTwoDecimalPlaces, StandardNumberFormat.ThousandsSeparatorTwoDecimalPlaces)]
        [InlineData(NumberFormats.TwoDecimalPlaces, StandardNumberFormat.TwoDecimalPlaces)]
        [InlineData("mm-dd-yy", StandardNumberFormat.ShortDate)]
        [InlineData("d-mmm-yy", StandardNumberFormat.LongDate)]
        [InlineData("d-mmm", StandardNumberFormat.DayMonth)]
        [InlineData("mmm-yy", StandardNumberFormat.MonthYear)]
        [InlineData("h:mm AM/PM", StandardNumberFormat.ShortTime12hour)]
        [InlineData("h:mm:ss AM/PM", StandardNumberFormat.LongTime12hour)]
        [InlineData("h:mm", StandardNumberFormat.ShortTime)]
        [InlineData("h:mm:ss", StandardNumberFormat.LongTime)]
        [InlineData("m/d/yy h:mm", StandardNumberFormat.DateAndTime)]
        [InlineData("#,##0 ;(#,##0)", StandardNumberFormat.NoDecimalPlacesNegativeParenthesis)]
        [InlineData("#,##0 ;[Red](#,##0)", StandardNumberFormat.NoDecimalPlacesNegativeParenthesisRed)]
        [InlineData("#,##0.00;(#,##0.00)", StandardNumberFormat.TwoDecimalPlacesNegativeParenthesis)]
        [InlineData("#,##0.00;[Red](#,##0.00)", StandardNumberFormat.TwoDecimalPlacesNegativeParenthesisRed)]
        [InlineData("mm:ss", StandardNumberFormat.MinutesAndSeconds)]
        [InlineData("[h]:mm:ss", StandardNumberFormat.Duration)]
        [InlineData("mmss.0", StandardNumberFormat.DecimalDuration)]
        [InlineData("##0.0E+0", StandardNumberFormat.Exponential)]
        [InlineData(NumberFormats.Text, StandardNumberFormat.Text)]
        public static void NumberFormatFromCustomStringBackwardCompatibility(string customString, StandardNumberFormat expectedStandardFormat)
        {
            var numberFormatExplict = NumberFormat.Standard(expectedStandardFormat);
#pragma warning disable CS0618 // Type or member is obsolete - Testing for backwards compatibilty
            var numberFormatLegacyFromStyle = (new Style { NumberFormat = customString }).Format;
            var numberFormatLegacyFromOptions = (new SpreadCheetahOptions { DefaultDateTimeNumberFormat = customString }).DefaultDateTimeFormat;
#pragma warning restore CS0618 // Type or member is obsolete

            Assert.Equal(numberFormatExplict, numberFormatLegacyFromStyle);
            Assert.Equal(numberFormatExplict, numberFormatLegacyFromOptions);
        }

        [Fact]
        public static void NumberFormatThrowsExceptionOnInvalidFormat()
        {
            Assert.ThrowsAny<ArgumentException>(() => NumberFormat.Custom(new string('a', 256))); // Strings longer than 255 characters should throw.
            Assert.ThrowsAny<ArgumentOutOfRangeException>(() => NumberFormat.Standard((StandardNumberFormat)(-1))); // Numbers that are not in the enum should throw.
        }
    }
}
