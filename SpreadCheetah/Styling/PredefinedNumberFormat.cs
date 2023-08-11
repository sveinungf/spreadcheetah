namespace SpreadCheetah.Styling
{
    /// <summary>
    /// Predefined number formats. These are interpreted by Excel in a locale dependent way, so the exact formatting may depend on the system locale settings when the file is viewed.
    /// </summary>
    public enum PredefinedNumberFormat
    {
        /// <summary>The default format.</summary>
        General = 0,
        /// <summary>Format as integer.</summary>
        NoDecimalPlaces = 1,
        /// <summary>Format as number with two decimal places.</summary>
        TwoDecimalPlaces = 2,
        /// <summary>Format as integer with thousands separator.</summary>
        ThousandsSeparator = 3,
        /// <summary>Format as number with two decimal places and thousands separator.</summary>
        ThousandsSeparatorTwoDecimalPlaces = 4,
        /// <summary>Format as integer with percent symbol.</summary>
        Percent = 9,
        /// <summary>Format as number with two decimal places and percent symbol.</summary>
        PercentTwoDecimalPlaces = 10,
        /// <summary>Format as a number between 1 and 10 multiplied by a power of 10.</summary>
        Scientific = 11,
        /// <summary>Format as integer with fraction using one digit denominator.</summary>
        Fraction = 12,
        /// <summary>Format as integer with fraction using two digit denominator.</summary>
        FractionTwoDenominatorPlaces = 13,
        /// <summary>Format as locale dependent Date with day, month and year.</summary>
        ShortDate = 14,
        /// <summary>Format as locale dependent Date with day, month name and year.</summary>
        LongDate = 15,
        /// <summary>Format as locale dependent Date with day and month name.</summary>
        DayMonth = 16,
        /// <summary>Format as locale dependent Date with month name and year.</summary>
        MonthYear = 17,
        /// <summary>Format as locale dependent time with hour, minute and AM/PM indicator.</summary>
        ShortTime12hour = 18,
        /// <summary>Format as locale dependent time with hour, minute, second and AM/PM indicator.</summary>
        LongTime12hour = 19,
        /// <summary>Format as locale dependent time with hour and minute.</summary>
        ShortTime = 20,
        /// <summary>Format as locale dependent time with hour, minute and second.</summary>
        LongTime = 21,
        /// <summary>Format as locale dependent date and timetime with day, month, year, hour and minute</summary>
        DateAndTime = 22,
        /// <summary>Format as integer with thousands separator, where negatives are in parenthesis ().</summary>
        NoDecimalPlacesNegativeParenthesis = 37,
        /// <summary>Format as integer with thousands separator, where negatives are red and in parenthesis ().</summary>
        NoDecimalPlacesNegativeParenthesisRed = 38,
        /// <summary>Format as number with two decimal places and thousands separator, where negatives are in parenthesis ().</summary>
        TwoDecimalPlacesNegativeParenthesis = 39,
        /// <summary>Format as number with two decimal places and thousands separator, where negatives are red and in parenthesis ().</summary>
        TwoDecimalPlacesNegativeParenthesisRed = 40,
        /// <summary>Format as locale dependent time with minute and second.</summary>
        MinutesAndSeconds = 45,
        /// <summary>Format as locale dependent duration with hours (if needed), minutes and seconds.</summary>
        Duration = 46,
        /// <summary>Format as locale dependent duration as a number of minutes and seconds, and one decemal place.</summary>
        DecimalDuration = 47,
        /// <summary>Format as a number between 1 and 10 multiplied by a power of 10.</summary>
        Exponential = 48,
        /// <summary>Format as text.</summary>
        Text = 49,
    }
}
