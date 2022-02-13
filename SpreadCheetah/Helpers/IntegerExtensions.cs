using System.Diagnostics;
using System.Globalization;

namespace SpreadCheetah.Helpers;

internal static class IntegerExtensions
{
    /// <summary>
    /// Gets the number of digits for positive integers.
    /// </summary>
    /// <param name="n"></param>
    public static int GetNumberOfDigits(this int n)
    {
        Debug.Assert(n >= 0);
        if (n < 10) return 1;
        if (n < 100) return 2;
        if (n < 1000) return 3;
        if (n < 10000) return 4;
        if (n < 100000) return 5;
        if (n < 1000000) return 6;
        if (n < 10000000) return 7;
        if (n < 100000000) return 8;
        if (n < 1000000000) return 9;
        return 10;
    }

    public static string ToStringInvariant(this int n) => n.ToString(CultureInfo.InvariantCulture);
}
