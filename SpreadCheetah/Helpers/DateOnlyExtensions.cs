using System.Globalization;

namespace SpreadCheetah.Helpers;

internal static class DateOnlyExtensions
{
    public static string ToStringInvariant(this DateOnly n) => n.ToString(CultureInfo.InvariantCulture);

    public static DateTime ToDateTime(this DateOnly d) => new(d.Year, d.Month, d.Day, 0,0,0, DateTimeKind.Unspecified);
}