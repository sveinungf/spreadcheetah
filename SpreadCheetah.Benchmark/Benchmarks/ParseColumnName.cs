using BenchmarkDotNet.Attributes;

namespace SpreadCheetah.Benchmark.Benchmarks;

[MemoryDiagnoser(false)]
public class ParseColumnName
{
    private string[] _names = null!;

    [GlobalSetup]
    public void GlobalSetup()
    {
        var random = new Random(42);
        _names = Enumerable.Range(0, 10000)
            .Select(_ => SpreadsheetUtility.GetColumnName(random.Next(1, 16384)))
            .ToArray();
    }

    [Benchmark]
    public int TryParseColumnName()
    {
        var sum = 0;
        for (var i = 0; i < _names.Length; i++)
        {
            if (SpreadsheetUtility.TryParseColumnName(_names[i].AsSpan(), out var number))
                sum += number;
        }
        return sum;
    }

    [Benchmark]
    public int TryParseColumnNameLoop()
    {
        var sum = 0;
        for (var i = 0; i < _names.Length; i++)
        {
            if (TryParseColumnNameLoop(_names[i].AsSpan(), out var number))
                sum += number;
        }
        return sum;
    }

    private static bool TryParseColumnNameLoop(ReadOnlySpan<char> columnName, out int columnNumber)
    {
        columnNumber = 0;
        if (columnName.Length is < 1 or > 3)
            return false;

        var pow = 1;
        for (int i = columnName.Length - 1; i >= 0; i--)
        {
            var letter = columnName[i];
            if (letter is < 'A' or > 'Z')
            {
                columnNumber = 0;
                return false;
            }

            columnNumber += (letter - 'A' + 1) * pow;
            pow *= 26;
        }

        if (columnNumber <= 16384)
            return true;

        columnNumber = 0;
        return false;
    }
}
