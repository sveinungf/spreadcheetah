using Bogus;
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

    public static string NumberFormat(this Randomizer randomizer)
    {
        string value;

        do
        {
            value = randomizer.String2(2, 255, "dhmsy0.,;-#%?/[] ");
        } while (string.Equals(value, "0%", StringComparison.Ordinal));

        return value;
    }

    public static double SimpleDouble(this Randomizer randomizer)
    {
        return randomizer.Number(int.MinValue, int.MaxValue) / 10000.0;
    }

    public static double SimpleDoublePair(this Randomizer randomizer, out double other)
    {
        var value1 = randomizer.Number(-500000, 500000);
        var value2 = randomizer.Number(value1, 500001);
        other = value2 / 1000.0;
        return value1 / 1000.0;
    }
}
