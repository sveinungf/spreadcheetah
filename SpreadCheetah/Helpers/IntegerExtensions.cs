using System.Globalization;

namespace SpreadCheetah.Helpers;

internal static class IntegerExtensions
{
    /// <summary>EMU = English Metric Unit</summary>
    public static int PixelsToEmu(this int n) => n * 9525;

    public static string ToStringInvariant(this int n) => n.ToString(CultureInfo.InvariantCulture);

    public static (int Width, int Height) Scale(this (int Width, int Height) dimensions, decimal scale)
    {
        var width = (int)Math.Round(dimensions.Width * scale, MidpointRounding.AwayFromZero);
        var height = (int)Math.Round(dimensions.Height * scale, MidpointRounding.AwayFromZero);
        return (width, height);
    }
}
