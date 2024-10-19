namespace SpreadCheetah.SourceGenerator.Models.Values;

/// <summary>
/// Values for this enum should be identical to SpreadCheetah.Styling.StandardNumberFormat.
/// </summary>
internal enum StandardNumberFormat
{
    General = 0,
    NoDecimalPlaces = 1,
    TwoDecimalPlaces = 2,
    ThousandsSeparator = 3,
    ThousandsSeparatorTwoDecimalPlaces = 4,
    Percent = 9,
    PercentTwoDecimalPlaces = 10,
    Scientific = 11,
    Fraction = 12,
    FractionTwoDenominatorPlaces = 13,
    ShortDate = 14,
    LongDate = 15,
    DayMonth = 16,
    MonthYear = 17,
    ShortTime12hour = 18,
    LongTime12hour = 19,
    ShortTime = 20,
    LongTime = 21,
    DateAndTime = 22,
    NoDecimalPlacesNegativeParenthesis = 37,
    NoDecimalPlacesNegativeParenthesisRed = 38,
    TwoDecimalPlacesNegativeParenthesis = 39,
    TwoDecimalPlacesNegativeParenthesisRed = 40,
    MinutesAndSeconds = 45,
    Duration = 46,
    DecimalDuration = 47,
    Exponential = 48,
    Text = 49
}
