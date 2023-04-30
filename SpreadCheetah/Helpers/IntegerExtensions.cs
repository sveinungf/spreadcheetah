using System.Globalization;

namespace SpreadCheetah.Helpers;

internal static class IntegerExtensions
{
    public static string ToStringInvariant(this int n) => n.ToString(CultureInfo.InvariantCulture);
}
