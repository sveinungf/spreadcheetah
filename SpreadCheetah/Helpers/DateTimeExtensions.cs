using System.Globalization;

namespace SpreadCheetah.Helpers;

internal static class DateTimeExtensions
{
    public static string ToStringInvariant(this DateTime n) => n.ToString(CultureInfo.InvariantCulture);

    public static DateOnly ToDateOnly(this DateTime d) => new DateOnly(d.Year, d.Month, d.Day);
}