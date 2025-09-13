using BenchmarkDotNet.Attributes;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.TestData;

namespace SpreadCheetah.Benchmark.Benchmarks;

[MemoryDiagnoser]
public class MultipleStyles
{
    private static readonly SpreadCheetahOptions Options = new() { DocumentProperties = null };
    private static readonly ICollection<Style> Styles = StyleGenerator.Generate(100000);

    [Benchmark]
    public async Task SpreadCheetah()
    {
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, Options);
        await spreadsheet.StartWorksheetAsync("a");

        foreach (var style in Styles)
        {
            spreadsheet.AddStyle(style);
        }

        await spreadsheet.FinishAsync();
    }
}
