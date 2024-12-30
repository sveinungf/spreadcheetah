using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using SpreadCheetah.Validations;

namespace SpreadCheetah.Benchmark.Benchmarks;

[SimpleJob(RuntimeMoniker.Net70)]
[MemoryDiagnoser]
public class ReusedDataValidation
{
    [Benchmark]
    public async Task SpreadCheetah()
    {
        var options = new SpreadCheetahOptions { DefaultDateTimeFormat = null };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, options);
        await spreadsheet.StartWorksheetAsync("Book1");
        var validation = DataValidation.IntegerGreaterThan(0);

        for (var i = 1; i <= 65534; ++i)
        {
            var reference = "A" + i;
            spreadsheet.TryAddDataValidation(reference, validation);
        }

        await spreadsheet.FinishAsync();
    }

    [Benchmark]
    public async Task SpreadCheetahDateTime()
    {
        var options = new SpreadCheetahOptions { DefaultDateTimeFormat = null };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, options);
        await spreadsheet.StartWorksheetAsync("Book1");
        var validation = DataValidation.DateTimeGreaterThan(new DateTime(2000,1,1,0,0,0,DateTimeKind.Unspecified));

        for (var i = 1 ; i <= 65534 ; ++i)
        {
            var reference = "A" + i;
            spreadsheet.TryAddDataValidation(reference, validation);
        }

        await spreadsheet.FinishAsync();
    }
}
