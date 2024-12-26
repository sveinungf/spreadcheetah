using System.Globalization;

namespace SpreadCheetah.Helpers;

internal static class DateTimeExtensions
{
    public static string ToStringInvariant(this DateTime n) => n.ToString(CultureInfo.InvariantCulture);

}