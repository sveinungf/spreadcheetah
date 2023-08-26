using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using System.Diagnostics.CodeAnalysis;

namespace SpreadCheetah.Benchmark.Benchmarks;

[SimpleJob(RuntimeMoniker.Net70)]
[MemoryDiagnoser(false)]
public class GetColumnName
{
    private int[] _numbers = null!;
    private string[] _names = null!;

    [GlobalSetup]
    public void GlobalSetup()
    {
        var random = new Random(42);
        _numbers = Enumerable.Range(0, 10000).Select(_ => random.Next(1, 16384)).ToArray();
    }

    [IterationSetup]
    public void IterationSetup()
    {
        _names = new string[_numbers.Length];
    }

    [Benchmark]
    public string[] GetColumnNameBaseline()
    {
        var names = _names;
        for (var i = 0; i < _numbers.Length; i++)
        {
            names[i] = SpreadsheetUtility.GetColumnName(_numbers[i]);
        }
        return names;
    }

    [Benchmark]
    public string[] GetColumnNameWithLoop()
    {
        var names = _names;
        for (var i = 0; i < _numbers.Length; i++)
        {
            names[i] = GetColumnNameWithLoop(_numbers[i]);
        }
        return names;
    }

    [Benchmark]
    public string[] GetColumnNameWithStringFormat()
    {
        var names = _names;
        for (var i = 0; i < _numbers.Length; i++)
        {
            names[i] = GetColumnNameWithStringFormat(_numbers[i]);
        }
        return names;
    }

    [Benchmark]
    public string[] GetColumnNameWithModulo()
    {
        var names = _names;
        for (var i = 0; i < _numbers.Length; i++)
        {
            names[i] = GetColumnNameWithModulo(_numbers[i]);
        }
        return names;
    }

    private const int MaxNumberOfColumns = 16384;

    [DoesNotReturn]
    public static void ThrowColumnNumberInvalid(string? paramName, int number) => throw new ArgumentOutOfRangeException(paramName, number, "The column number must be greater than 0 and can't be larger than 16384.");

    private static string GetColumnNameWithLoop(int columnNumber)
    {
        if (columnNumber < 1 || columnNumber > MaxNumberOfColumns)
            ThrowColumnNumberInvalid(nameof(columnNumber), columnNumber);

        string columnName = "";

        while (columnNumber > 0)
        {
            int modulo = (columnNumber - 1) % 26;
#pragma warning disable S1643 // Strings should not be concatenated using '+' in a loop
            columnName = Convert.ToChar('A' + modulo) + columnName;
#pragma warning restore S1643 // Strings should not be concatenated using '+' in a loop
            columnNumber = (columnNumber - modulo) / 26;
        }

        return columnName;
    }

    public static string GetColumnNameWithStringFormat(int columnNumber)
    {
        if (columnNumber < 1 || columnNumber > MaxNumberOfColumns)
            ThrowColumnNumberInvalid(nameof(columnNumber), columnNumber);

        if (columnNumber <= 26)
            return ((char)(columnNumber + 'A' - 1)).ToString();

        if (columnNumber <= 702)
        {
            var quotient = Math.DivRem(columnNumber - 1, 26, out var remainder);
            char firstChar = (char)('A' - 1 + quotient);
            char secondChar = (char)('A' + remainder);
            return string.Format("{0}{1}", firstChar, secondChar);
        }
        else
        {
            var quotient1 = Math.DivRem(columnNumber - 1, 26, out var remainder1);
            var quotient2 = Math.DivRem(quotient1 - 1, 26, out var remainder2);
            char firstChar = (char)('A' - 1 + quotient2);
            char secondChar = (char)('A' + remainder2);
            char thirdChar = (char)('A' + remainder1);
            return string.Format("{0}{1}{2}", firstChar, secondChar, thirdChar);
        }
    }

    public static string GetColumnNameWithModulo(int columnNumber)
    {
        if (columnNumber < 1 || columnNumber > MaxNumberOfColumns)
            ThrowColumnNumberInvalid(nameof(columnNumber), columnNumber);

        if (columnNumber <= 26)
            return ((char)(columnNumber + 'A' - 1)).ToString();

        if (columnNumber <= 702)
        {
            Span<char> characters = stackalloc char[2];
            var quotient = (columnNumber - 1) / 26;
            var remainder = (columnNumber - 1) % 26;
            characters[0] = (char)('A' - 1 + quotient);
            characters[1] = (char)('A' + remainder);
            return characters.ToString();
        }
        else
        {
            Span<char> characters = stackalloc char[3];
            var quotient1 = (columnNumber - 1) / 26;
            var remainder1 = (columnNumber - 1) % 26;
            var quotient2 = (quotient1 - 1) / 26;
            var remainder2 = (quotient1 - 1) % 26;
            characters[0] = (char)('A' - 1 + quotient2);
            characters[1] = (char)('A' + remainder2);
            characters[2] = (char)('A' + remainder1);
            return characters.ToString();
        }
    }
}
