using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;

namespace SpreadCheetah.Benchmark.Benchmarks;

[SimpleJob(RuntimeMoniker.Net70)]
[SimpleJob(RuntimeMoniker.Net80)]
[MemoryDiagnoser]
public class MinimalSpreadsheet
{
    private static readonly SpreadCheetahOptions Options = new() { DefaultDateTimeFormat = null };

    [Benchmark]
    public async Task SpreadCheetah()
    {
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, Options);
        await spreadsheet.StartWorksheetAsync("a");
        await spreadsheet.FinishAsync();
    }
}
