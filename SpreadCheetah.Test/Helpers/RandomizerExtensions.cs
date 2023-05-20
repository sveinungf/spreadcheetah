using Bogus;
using System.Drawing;

namespace SpreadCheetah.Test.Helpers;

internal static class RandomizerExtensions
{
    public static Color Color(this Randomizer randomizer)
    {
        return System.Drawing.Color.FromArgb(randomizer.Number(int.MaxValue));
    }

    public static string NumberFormat(this Randomizer randomizer)
    {
        string value;

        do
        {
            value = randomizer.String2(2, 255, "dhmsy0.,;-#%?/[] ");
        } while (value == "0%");

        return value;
    }
}
