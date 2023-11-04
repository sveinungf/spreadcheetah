using System.Globalization;

namespace SpreadCheetah.Helpers;

internal static class IntegerExtensions
{
    /// <summary>EMU = English Metric Unit</summary>
    public static int PixelsToEmu(this int n) => n * 9525;

    public static string ToStringInvariant(this int n) => n.ToString(CultureInfo.InvariantCulture);
}
