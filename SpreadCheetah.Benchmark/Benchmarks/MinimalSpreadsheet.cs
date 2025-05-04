using BenchmarkDotNet.Attributes;

namespace SpreadCheetah.Benchmark.Benchmarks;

[MemoryDiagnoser]
public class MinimalSpreadsheet
{
    private static readonly SpreadCheetahOptions Options = new() { DefaultDateTimeFormat = null, DocumentProperties = null };

    [Benchmark]
    public async Task SpreadCheetah()
    {
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, Options);
        await spreadsheet.StartWorksheetAsync("a");
        await spreadsheet.FinishAsync();
    }
}
