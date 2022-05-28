using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using SpreadCheetah.Benchmark.Helpers;

namespace SpreadCheetah.Benchmark.Benchmarks;

[SimpleJob(RuntimeMoniker.Net60)]
[MemoryDiagnoser]
public class AddObjectAsRow
{
    public IList<Student> Students { get; private set; } = null!;

    [Params(10, 1000, 100000)]
    public int NumberOfRows { get; set; }

    [GlobalSetup]
    public void GlobalSetup()
    {
        Students = StudentGenerator.Generate(NumberOfRows);
    }

    [Benchmark]
    public async Task SpreadCheetahAddAsRow()
    {
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Book1");
        var ctx = StudentContext.Default.Student;
        var students = Students;

        for (var i = 0; i < students.Count; ++i)
        {
            await spreadsheet.AddAsRowAsync(students[i], ctx);
        }

        await spreadsheet.FinishAsync();
    }

    [Benchmark]
    public async Task SpreadCheetahAddRangeAsRows()
    {
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Book1");
        var ctx = StudentContext.Default.Student;
        await spreadsheet.AddRangeAsRowsAsync(Students, ctx);
        await spreadsheet.FinishAsync();
    }
}
