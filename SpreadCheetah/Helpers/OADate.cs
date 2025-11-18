using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace SpreadCheetah.Helpers;

internal readonly record struct OADate(long Ticks)
{
    // Implementation is based on DateTime.ToOADate(). These constants are taken from there.
    private const int DaysPerYear = 365;
    private const int DaysPer4Years = DaysPerYear * 4 + 1;
    private const int DaysPer100Years = DaysPer4Years * 25 - 1;
    private const int DaysPer400Years = DaysPer100Years * 4 + 1;
    private const int DaysTo1899 = DaysPer400Years * 4 + DaysPer100Years * 3 - 367; // Number of days from 1/1/0001 to 12/30/1899
    private const long DoubleDateOffset = DaysTo1899 * TimeSpan.TicksPerDay;
    private const long MinTicks = (DaysPer100Years - DaysPerYear) * TimeSpan.TicksPerDay;

    public static void EnsureValidTicks(long ticks)
    {
        if (ticks is >= TimeSpan.TicksPerDay and < MinTicks)
            ThrowHelper.InvalidOADate();
    }

    public bool TryFormat(Span<byte> destination, out int bytesWritten)
    {
        // Days can be up to 7 digits (max = 2958465, min = -657434).
        // In this implementation, the fraction part is limited to 11 digits.
        if (destination.Length < 19)
        {
            bytesWritten = 0;
            return false;
        }

        var value = Ticks;
        return value >= TimeSpan.TicksPerDay
            ? TryFormatCore(value, destination, out bytesWritten)
            : TryFormatEdgeCases(value, destination, out bytesWritten);
    }

    private static bool TryFormatCore(long value, Span<byte> destination, out int bytesWritten)
    {
        Debug.Assert(value >= MinTicks);

        var days = Math.DivRem(value, TimeSpan.TicksPerDay, out var ticksAfterMidnight);
        if (days < DaysTo1899 + 31 + 29 + 1) // Subtract 1 day if day is less than 03/01/1900
        {
            days -= 1;
        }
        TryFormatLong(days - DaysTo1899, destination, out bytesWritten);

        if (ticksAfterMidnight != 0)
        {
            var fractionDestination = destination.Slice(bytesWritten);
            bytesWritten += FormatFraction((ulong)ticksAfterMidnight, fractionDestination);
        }

        return true;
    }

    private static bool TryFormatEdgeCases(long value, Span<byte> destination, out int bytesWritten)
    {
        if (value == 0)
        {
            destination[0] = (byte)'0';
            bytesWritten = 1;
            return true;
        }

        if (value < TimeSpan.TicksPerDay)
            value += DoubleDateOffset;

        return TryFormatCore(value, destination, out bytesWritten);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool TryFormatLong(long value, Span<byte> destination, out int bytesWritten)
    {
#if NET8_0_OR_GREATER
        return value.TryFormat(destination, out bytesWritten, provider: System.Globalization.NumberFormatInfo.InvariantInfo);
#else
        return System.Buffers.Text.Utf8Formatter.TryFormat(value, destination, out bytesWritten);
#endif
    }

#pragma warning disable MA0051 // Method is too long
    private static int FormatFraction(ulong ticksAfterMidnight, Span<byte> destination)
#pragma warning restore MA0051 // Method is too long
    {
        destination[0] = (byte)'.';

        var fraction = ticksAfterMidnight * 100 / 864;
        var (quotient, remainder) = Math.DivRem(fraction, 10000000000);
        destination[1] = (byte)(quotient + '0');
        if (remainder == 0) return 2;

        (quotient, remainder) = Math.DivRem(remainder, 1000000000);
        destination[2] = (byte)(quotient + '0');
        if (remainder == 0) return 3;

        (quotient, remainder) = Math.DivRem(remainder, 100000000);
        destination[3] = (byte)(quotient + '0');
        if (remainder == 0) return 4;

        (quotient, remainder) = Math.DivRem(remainder, 10000000);
        destination[4] = (byte)(quotient + '0');
        if (remainder == 0) return 5;

        (quotient, remainder) = Math.DivRem(remainder, 1000000);
        destination[5] = (byte)(quotient + '0');
        if (remainder == 0) return 6;

        (quotient, remainder) = Math.DivRem(remainder, 100000);
        destination[6] = (byte)(quotient + '0');
        if (remainder == 0) return 7;

        (quotient, remainder) = Math.DivRem(remainder, 10000);
        destination[7] = (byte)(quotient + '0');
        if (remainder == 0) return 8;

        (quotient, remainder) = Math.DivRem(remainder, 1000);
        destination[8] = (byte)(quotient + '0');
        if (remainder == 0) return 9;

        (quotient, remainder) = Math.DivRem(remainder, 100);
        destination[9] = (byte)(quotient + '0');
        if (remainder == 0) return 10;

        (quotient, remainder) = Math.DivRem(remainder, 10);
        destination[10] = (byte)(quotient + '0');
        if (remainder == 0) return 11;

        destination[11] = (byte)(remainder + '0');
        return 12;
    }
}
