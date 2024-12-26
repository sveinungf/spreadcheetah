using Bogus;
using SpreadCheetah.Styling;
using System.Drawing;

namespace SpreadCheetah.Test.Helpers;

internal static class RandomizerExtensions
{
    public static string CellRange(this Randomizer randomizer)
    {
        var rowNo1 = randomizer.Number(1, 1048575);
        var rowNo2 = randomizer.Number(rowNo1, 1048576);

        var columnNo1 = randomizer.Number(1, 16383);
        var columnNo2 = randomizer.Number(columnNo1, 16384);

        var columnName1 = SpreadsheetUtility.GetColumnName(columnNo1);
        var columnName2 = SpreadsheetUtility.GetColumnName(columnNo2);

        return $"{columnName1}{rowNo1}:{columnName2}{rowNo2}";
    }

    public static Color Color(this Randomizer randomizer)
    {
        return System.Drawing.Color.FromArgb(randomizer.Number(int.MaxValue));
    }

    public static int IntPair(this Randomizer randomizer, out int other)
    {
        var value1 = randomizer.Int();
        other = randomizer.Int(value1);
        return value1;
    }

    public static NumberFormat NumberFormat(this Randomizer randomizer)
    {
        string value;

        do
        {
            value = randomizer.String2(2, 255, "dhmsy0.,;-#%?/[] ");
        } while (string.Equals(value, "0%", StringComparison.Ordinal));

        return Styling.NumberFormat.Custom(value);
    }

    public static double SimpleDouble(this Randomizer randomizer)
    {
        return randomizer.Number(int.MinValue, int.MaxValue) / 10000.0;
    }

    public static double SimpleDoublePair(this Randomizer randomizer, out double other)
    {
        var value1 = randomizer.Number(-500000, 500000);
        var value2 = randomizer.Number(value1, 500001);

        // other must be necessarily greater than value1.
        other = value2 / 1000.0;
        return value1 / 1000.0;
    }

    public static DateTime SimpleDateTime(this Randomizer randomizer)
    {
        var minYear = 1950;
        var maxYear = 3200;

        var minMonth = 1;
        var maxMonth = 12;

        var minDay = 1;
        var maxDay = 28;

        var minHour = 1;
        var maxHour = 12;

        var minMinute = 1;
        var maxMinute = 59;

        var minSecond = 1;
        var maxSecond = 59;

        var randomYear = randomizer.Number(minYear, maxYear);
        var randomMonth = randomizer.Number(minMonth, maxMonth);
        var randomDay = randomizer.Number(minDay, maxDay);
        var randomHour = randomizer.Number(minHour, maxHour);
        var randomMinute = randomizer.Number(minMinute, maxMinute);
        var randomSecond = randomizer.Number(minSecond, maxSecond);

        var dateTime = new DateTime(randomYear,
                                    randomMonth,
                                    randomDay,
                                    randomHour,
                                    randomMinute,
                                    randomSecond,
                                    DateTimeKind.Unspecified);

        return dateTime;
    }

    public static DateTime SimpleDateTimePair(this Randomizer randomizer, out DateTime other)
    {
        var minYear = 1950;
        var maxYear = 3200;

        var minMonth = 1;
        var maxMonth = 12;

        var minDay = 1;
        var maxDay = 28;

        var minHour = 1;
        var maxHour = 12;

        var minMinute = 1;
        var maxMinute = 59;

        var minSecond = 1;
        var maxSecond = 59;

        var randomYear = randomizer.Number(minYear, maxYear);
        var randomMonth = randomizer.Number(minMonth, maxMonth);
        var randomDay = randomizer.Number(minDay, maxDay);
        var randomHour = randomizer.Number(minHour, maxHour);
        var randomMinute = randomizer.Number(minMinute, maxMinute);
        var randomSecond = randomizer.Number(minSecond, maxSecond);

        var otherDateTime = new DateTime(randomYear,
                                         randomMonth,
                                         randomDay,
                                         randomHour,
                                         randomMinute,
                                         randomSecond,
                                         DateTimeKind.Unspecified);

        var upperlimit = otherDateTime.AddSeconds(-4).AddDays(-1);

        randomYear = randomizer.Number(minYear, upperlimit.Year);
        randomMonth = randomizer.Number(minMonth, upperlimit.Month);
        randomDay = randomizer.Number(minDay, upperlimit.Day);
        randomHour = randomizer.Number(minHour, upperlimit.Hour);
        randomMinute = randomizer.Number(minMinute, upperlimit.Minute);
        randomSecond = randomizer.Number(minSecond, upperlimit.Second);

        var newDateTime = new DateTime(randomYear,
                                       randomMonth,
                                       randomDay,
                                       randomHour,
                                       randomMinute,
                                       randomSecond,
                                       DateTimeKind.Unspecified);

        // other must be necessarily greater than the returned value.
        other = otherDateTime;
        return newDateTime;
    }
}
