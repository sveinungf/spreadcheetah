using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;

namespace SpreadCheetah.Benchmark.Benchmarks;

[SimpleJob(RuntimeMoniker.Net90)]
[MemoryDiagnoser]
public class MergedCells
{
    [Benchmark]
    public async Task SpreadCheetah()
    {
        var options = new SpreadCheetahOptions { DefaultDateTimeFormat = null };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, options);
        await spreadsheet.StartWorksheetAsync("Book1");

        for (var i = 1; i <= 65534; ++i)
        {
            spreadsheet.MergeCells($"A{i}:B{i}");
        }

        await spreadsheet.FinishAsync();
    }
}
