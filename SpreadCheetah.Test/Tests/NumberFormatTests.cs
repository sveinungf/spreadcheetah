using SpreadCheetah.Styling;
using Xunit;

namespace SpreadCheetah.Test.Tests
{
    public static class NumberFormatTests
    {
        [Theory]
        [InlineData(NumberFormats.General, PredefinedNumberFormat.General)]
        [InlineData(NumberFormats.Fraction, PredefinedNumberFormat.Fraction)]
        [InlineData(NumberFormats.FractionTwoDenominatorPlaces, PredefinedNumberFormat.FractionTwoDenominatorPlaces)]
        [InlineData(NumberFormats.NoDecimalPlaces, PredefinedNumberFormat.NoDecimalPlaces)]
        [InlineData(NumberFormats.Percent, PredefinedNumberFormat.Percent)]
        [InlineData(NumberFormats.PercentTwoDecimalPlaces, PredefinedNumberFormat.PercentTwoDecimalPlaces)]
        [InlineData(NumberFormats.Scientific, PredefinedNumberFormat.Scientific)]
        [InlineData(NumberFormats.ThousandsSeparator, PredefinedNumberFormat.ThousandsSeparator)]
        [InlineData(NumberFormats.ThousandsSeparatorTwoDecimalPlaces, PredefinedNumberFormat.ThousandsSeparatorTwoDecimalPlaces)]
        [InlineData(NumberFormats.TwoDecimalPlaces, PredefinedNumberFormat.TwoDecimalPlaces)]
        [InlineData("mm-dd-yy", PredefinedNumberFormat.ShortDate)]
        [InlineData("d-mmm-yy", PredefinedNumberFormat.LongDate)]
        [InlineData("d-mmm", PredefinedNumberFormat.DayMonth)]
        [InlineData("mmm-yy", PredefinedNumberFormat.MonthYear)]
        [InlineData("h:mm AM/PM", PredefinedNumberFormat.ShortTime12hour)]
        [InlineData("h:mm:ss AM/PM", PredefinedNumberFormat.LongTime12hour)]
        [InlineData("h:mm", PredefinedNumberFormat.ShortTime)]
        [InlineData("h:mm:ss", PredefinedNumberFormat.LongTime)]
        [InlineData("m/d/yy h:mm", PredefinedNumberFormat.DateAndTime)]
        [InlineData("#,##0 ;(#,##0)", PredefinedNumberFormat.NoDecimalPlacesNegativeParenthesis)]
        [InlineData("#,##0 ;[Red](#,##0)", PredefinedNumberFormat.NoDecimalPlacesNegativeParenthesisRed)]
        [InlineData("#,##0.00;(#,##0.00)", PredefinedNumberFormat.TwoDecimalPlacesNegativeParenthesis)]
        [InlineData("#,##0.00;[Red](#,##0.00)", PredefinedNumberFormat.TwoDecimalPlacesNegativeParenthesisRed)]
        [InlineData("mm:ss", PredefinedNumberFormat.MinutesAndSeconds)]
        [InlineData("[h]:mm:ss", PredefinedNumberFormat.Duration)]
        [InlineData("mmss.0", PredefinedNumberFormat.DecimalDuration)]
        [InlineData("##0.0E+0", PredefinedNumberFormat.Exponential)]
        [InlineData(NumberFormats.Text, PredefinedNumberFormat.Text)]
        public static void NumberFormatFromCustomStringBackwardCompatibility(string customString, PredefinedNumberFormat expectedPredefinedFormat)
        {
            var numberFormatExplict = NumberFormat.Predefined(expectedPredefinedFormat);
#pragma warning disable CS0618 // Type or member is obsolete - Testing for backwards compatibilty
            var numberFormatLegacyFromStyle = (new Style { NumberFormat = customString }).Format;
            var numberFormatLegacyFromOptions = (new SpreadCheetahOptions { DefaultDateTimeNumberFormat = customString }).DefaultDateTimeFormat;
#pragma warning restore CS0618 // Type or member is obsolete

            Assert.Equal(numberFormatExplict, numberFormatLegacyFromStyle);
            Assert.Equal(numberFormatExplict, numberFormatLegacyFromOptions);
        }
    }
}
