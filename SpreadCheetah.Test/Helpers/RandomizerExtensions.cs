using Bogus;
using System.Drawing;

namespace SpreadCheetah.Test.Helpers;

internal static class RandomizerExtensions
{
    public static Color Color(this Randomizer randomizer)
    {
        return System.Drawing.Color.FromArgb(randomizer.Number(int.MaxValue));
    }
}
