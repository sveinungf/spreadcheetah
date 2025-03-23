using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using Polyfills;
using SpreadCheetah.Helpers;
using System.Buffers.Text;
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
        var origin = new DateTime(2025, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        _dateTimes = Enumerable.Range(0, Count)
            .Select(_ => WithFractions
                    ? origin.AddSeconds(random.Next(-10000000, 10000000))
                    : origin.AddDays(random.Next(-10000, 10000)))
            .ToList();
    }

    [Benchmark]
    public IList<string> DateTime_ToOADate()
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
    public IList<string> OADate_TryFormat()
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
