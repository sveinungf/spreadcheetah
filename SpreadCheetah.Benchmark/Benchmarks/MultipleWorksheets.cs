using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using SpreadCheetah.Worksheets;

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
        var worksheetOptions = new WorksheetOptions();
        worksheetOptions.Column(1).Width = 100;
        worksheetOptions.Column(2).Hidden = false;

        for (var i = 0; i < 1000; ++i)
        {
            if (i % 2 == 0)
            {
                await spreadsheet.StartWorksheetAsync(i.ToString());
                continue;
            }

            worksheetOptions.FrozenColumns = i % 4 == 0 ? 1 : null;
            worksheetOptions.FrozenRows = i % 8 == 0 ? 1 : null;
            worksheetOptions.Visibility = i % 16 == 0 ? WorksheetVisibility.Hidden : WorksheetVisibility.Visible;
            worksheetOptions.Column(2).Hidden = i % 32 == 0;

            await spreadsheet.StartWorksheetAsync(i.ToString(), worksheetOptions);
        }

        await spreadsheet.FinishAsync();
    }
}
