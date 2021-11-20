using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using SpreadCheetah.Helpers;

namespace SpreadCheetah.Benchmark.Benchmarks;

[SimpleJob(RuntimeMoniker.Net48)]
[SimpleJob(RuntimeMoniker.Net50)]
[MemoryDiagnoser]
public class GetNumberOfDigits
{
    private readonly List<int> _values = new();

    [Params(1000, int.MaxValue)]
    public int MaxValue { get; set; }

    [GlobalSetup]
    public void GlobalSetup()
    {
        var r = new Random();

        for (var i = 0; i < 10000; ++i)
        {
            _values.Add(r.Next(1, MaxValue));
        }
    }

    [Benchmark]
    public int Branching()
    {
        var sum = 0;
        for (var i = 0; i < _values.Count; ++i)
            sum += _values[i].GetNumberOfDigits();

        return sum;
    }
}
