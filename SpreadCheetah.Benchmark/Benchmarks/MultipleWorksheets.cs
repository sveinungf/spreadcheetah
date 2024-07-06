using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;

namespace SpreadCheetah.Benchmark.Benchmarks;

[SimpleJob(RuntimeMoniker.Net80)]
[MemoryDiagnoser]
public class MultipleWorksheets
{
    private static readonly SpreadCheetahOptions Options = new() { DefaultDateTimeFormat = null };

    [Benchmark]
    public async Task SpreadCheetah()
    {
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, Options);
        
        for (var i = 0; i < 1000; ++i)
        {
            await spreadsheet.StartWorksheetAsync(i.ToString());
        }


        await spreadsheet.FinishAsync();
    }
}
