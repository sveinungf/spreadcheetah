using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace SpreadCheetah.Helpers;

internal readonly struct OADate(long ticks)
{
    // Implementation is based on DateTime.ToOADate(). These constants are taken from there.
    private const int DaysPerYear = 365;
    private const int DaysPer4Years = DaysPerYear * 4 + 1;
    private const int DaysPer100Years = DaysPer4Years * 25 - 1;
    private const int DaysPer400Years = DaysPer100Years * 4 + 1;
    private const int DaysTo1899 = DaysPer400Years * 4 + DaysPer100Years * 3 - 367;
    private const long DoubleDateOffset = DaysTo1899 * TimeSpan.TicksPerDay;
    private const long MillisecondsPerDay = TimeSpan.TicksPerDay / TimeSpan.TicksPerMillisecond;
    private const long OADateMinAsTicks = (DaysPer100Years - DaysPerYear) * TimeSpan.TicksPerDay;

    public bool TryFormat(Span<byte> destination, out int bytesWritten)
    {
        bytesWritten = 0;

        // Days can be up to 7 digits (max = 2958465, min = -657434).
        // In this implementation, the fraction part is limited to 11 digits.
        if (destination.Length < 19)
            return false;

        var value = ticks;
        if (value == 0)
        {
            destination[0] = (byte)'0';
            bytesWritten = 1;
            return true;
        }

        if (value < TimeSpan.TicksPerDay)
            value += DoubleDateOffset;

        // TODO: Check in DataCell constructor and throw if below OADateMinAsTicks
        Debug.Assert(value >= OADateMinAsTicks);

        var millis = (value - DoubleDateOffset) / TimeSpan.TicksPerMillisecond;
        var days = Math.DivRem(millis, MillisecondsPerDay, out var millisAfterMidnight);

        if (millisAfterMidnight == 0)
            return TryFormatDays(days, destination, out bytesWritten);

        return TryFormatWithFraction(days, millisAfterMidnight, destination, out bytesWritten);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool TryFormatDays(long days, Span<byte> destination, out int bytesWritten)
    {
#if NET8_0_OR_GREATER
        return days.TryFormat(destination, out bytesWritten, provider: System.Globalization.NumberFormatInfo.InvariantInfo);
#else
        return System.Buffers.Text.Utf8Formatter.TryFormat(days, destination, out bytesWritten);
#endif
    }

#pragma warning disable MA0051 // Method is too long
    private static bool TryFormatWithFraction(long days, long millisAfterMidnight, Span<byte> destination, out int bytesWritten)
#pragma warning restore MA0051 // Method is too long
    {
        var fraction = millisAfterMidnight * 1000000 / 864;
        if (fraction < 0)
        {
            days--;
            fraction += 100000000000;
        }

        TryFormatDays(days, destination, out bytesWritten);
        destination[bytesWritten] = (byte)'.';
        bytesWritten++;

        var quotient = Math.DivRem(fraction, 10000000000, out var remainder);
        destination[bytesWritten] = (byte)(quotient + '0');
        bytesWritten++;
        if (remainder == 0) return true;

        quotient = Math.DivRem(remainder, 1000000000, out remainder);
        destination[bytesWritten] = (byte)(quotient + '0');
        bytesWritten++;
        if (remainder == 0) return true;

        quotient = Math.DivRem(remainder, 100000000, out remainder);
        destination[bytesWritten] = (byte)(quotient + '0');
        bytesWritten++;
        if (remainder == 0) return true;

        quotient = Math.DivRem(remainder, 10000000, out remainder);
        destination[bytesWritten] = (byte)(quotient + '0');
        bytesWritten++;
        if (remainder == 0) return true;

        quotient = Math.DivRem(remainder, 1000000, out remainder);
        destination[bytesWritten] = (byte)(quotient + '0');
        bytesWritten++;
        if (remainder == 0) return true;

        quotient = Math.DivRem(remainder, 100000, out remainder);
        destination[bytesWritten] = (byte)(quotient + '0');
        bytesWritten++;
        if (remainder == 0) return true;

        quotient = Math.DivRem(remainder, 10000, out remainder);
        destination[bytesWritten] = (byte)(quotient + '0');
        bytesWritten++;
        if (remainder == 0) return true;

        quotient = Math.DivRem(remainder, 1000, out remainder);
        destination[bytesWritten] = (byte)(quotient + '0');
        bytesWritten++;
        if (remainder == 0) return true;

        quotient = Math.DivRem(remainder, 100, out remainder);
        destination[bytesWritten] = (byte)(quotient + '0');
        bytesWritten++;
        if (remainder == 0) return true;

        quotient = Math.DivRem(remainder, 10, out remainder);
        destination[bytesWritten] = (byte)(quotient + '0');
        bytesWritten++;
        if (remainder == 0) return true;

        destination[bytesWritten] = (byte)(remainder + '0');
        bytesWritten++;
        return true;
    }
}
