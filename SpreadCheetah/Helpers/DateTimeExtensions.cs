using System.Diagnostics;
using System.Text;

namespace SpreadCheetah.Helpers;

internal static class DateTimeExtensions
{
    public static string ToOADateString(this DateTime dateTime)
    {
        OADate.EnsureValidTicks(dateTime.Ticks);
        var oaDate = new OADate(dateTime.Ticks);
        Span<byte> destination = stackalloc byte[19];
        var success = oaDate.TryFormat(destination, out var bytesWritten);
        Debug.Assert(success);
        return Encoding.UTF8.GetString(destination.Slice(0, bytesWritten));
    }
}
