namespace SpreadCheetah.Helpers;

internal static class ValueConstants
{
    // https://stackoverflow.com/questions/1701055/what-is-the-maximum-length-in-chars-needed-to-represent-any-double-value
    public const int DoubleValueMaxCharacters = 24;
    public const int FloatValueMaxCharacters = 16;

    // -2147483648 (int.MinValue)
    public const int IntegerValueMaxCharacters = 11;
}
