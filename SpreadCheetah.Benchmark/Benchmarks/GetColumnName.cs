using BenchmarkDotNet.Attributes;

namespace SpreadCheetah.Benchmark.Benchmarks;

[MemoryDiagnoser(false)]
public class GetColumnName
{
    private int[] _numbers1 = null!;
    private int[] _numbers2 = null!;
    private int[] _numbers3 = null!;
    private string[] _names = null!;

    [Params(1, 2, 3)]
    public int NumberOfCharacters { get; set; }

    [GlobalSetup]
    public void GlobalSetup()
    {
        var random = new Random(42);
        _numbers1 = Enumerable.Range(0, 1000000).Select(_ => random.Next(1, 27)).ToArray();
        _numbers2 = Enumerable.Range(0, 1000000).Select(_ => random.Next(27, 703)).ToArray();
        _numbers3 = Enumerable.Range(0, 1000000).Select(_ => random.Next(703, 16384)).ToArray();
    }

    [IterationSetup]
    public void IterationSetup()
    {
        _names = new string[_numbers1.Length];
    }

    [Benchmark(Baseline = true)]
    public string[] GetColumnNameBaseline()
    {
        var numbers = NumberOfCharacters switch
        {
            1 => _numbers1,
            2 => _numbers2,
            3 => _numbers3,
            _ => throw new InvalidOperationException()
        };

        var names = _names;
        for (var i = 0; i < numbers.Length; i++)
        {
            names[i] = SpreadsheetUtility.GetColumnName(numbers[i]);
        }
        return names;
    }
}
