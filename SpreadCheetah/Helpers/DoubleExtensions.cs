using System.Globalization;

namespace SpreadCheetah.Helpers;

internal static class DoubleExtensions
{
    public static string ToStringInvariant(this double n) => n.ToString(CultureInfo.InvariantCulture);
}
