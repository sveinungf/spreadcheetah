using Bogus;
using SpreadCheetah.Styling;
using System.Drawing;

namespace SpreadCheetah.Benchmark.Helpers;

internal static class RandomizerExtensions
{
    public static Color Color(this Randomizer randomizer)
    {
        return System.Drawing.Color.FromArgb(randomizer.Number(int.MaxValue));
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
}
