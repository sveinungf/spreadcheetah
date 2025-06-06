using BenchmarkDotNet.Attributes;

namespace SpreadCheetah.Benchmark.Benchmarks;

[MemoryDiagnoser]
public class Notes
{
    private static readonly SpreadCheetahOptions Options = new() { DefaultDateTimeFormat = null };

    [Benchmark]
    public async Task SpreadCheetah()
    {
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, Options);
        await spreadsheet.StartWorksheetAsync("Book1");

        for (var i = 1; i <= 32768; ++i)
        {
            spreadsheet.AddNote($"A{i}", "My note");
        }

        await spreadsheet.FinishAsync();
    }
}
