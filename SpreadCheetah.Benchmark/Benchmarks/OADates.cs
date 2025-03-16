using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using System.Buffers.Text;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text;

namespace SpreadCheetah.Benchmark.Benchmarks;

[SimpleJob(RuntimeMoniker.Net90)]
[MemoryDiagnoser]
public class OADates
{
    private List<DateTime> _dateTimes = [];

    [Params(10000)]
    public int Count { get; set; }

    [Params(true, false)]
    public bool WithFractions { get; set; }

    [GlobalSetup]
    public void GlobalSetup()
    {
        var random = new Random(42);
        var origin = new DateTime(2025, 1, 1);
        _dateTimes = Enumerable.Range(0, Count)
            .Select(_ =>
            {
                return WithFractions
                    ? origin.AddSeconds(random.Next(-10000000, 10000000))
                    : origin.AddDays(random.Next(-10000, 10000));
            })
            .ToList();
    }

    [Benchmark]
    public List<string> DateTime_ToOADate()
    {
        var result = new List<string>(Count);
        Span<byte> destination = stackalloc byte[19];

        foreach (var dateTime in _dateTimes)
        {
            var oaDate = dateTime.ToOADate();
            Utf8Formatter.TryFormat(oaDate, destination, out var written);
            var stringValue = Encoding.UTF8.GetString(destination.Slice(0, written));
            result.Add(stringValue);
        }

        return result;
    }

    [Benchmark]
    public List<string> OADate_TryFormat()
    {
        var result = new List<string>(Count);
        Span<byte> destination = stackalloc byte[19];

        foreach (var dateTime in _dateTimes)
        {
            var oaDate = new OADate(dateTime.Ticks);
            oaDate.TryFormat(destination, out var written);
            var stringValue = Encoding.UTF8.GetString(destination.Slice(0, written));
            result.Add(stringValue);
        }

        return result;
    }
}

/// <summary>
/// Copy of SpreadCheetah.Helpers.OADate
/// </summary>
file readonly record struct OADate(long Ticks)
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

        var value = Ticks;
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
