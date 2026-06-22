using BenchmarkDotNet.Attributes;
using SpreadCheetah.Benchmark.Helpers;

namespace SpreadCheetah.Benchmark.Benchmarks;

[MemoryDiagnoser]
public class DateTimeCells
{
    [Params(10)]
    public int NumberOfColumns { get; set; }

    [Params(500000)]
    public int NumberOfRows { get; set; }

    [Params(true, false)]
    public bool WithFractions { get; set; }

    private List<List<DateTime>> _rows = null!;

    [GlobalSetup]
    public void GlobalSetup()
    {
        var random = new Random(42);

        _rows = Enumerable.Range(0, NumberOfRows)
            .Select(_ => Enumerable.Range(0, NumberOfColumns)
                .Select(_ => random.NextDateTime(WithFractions))
                .ToList())
            .ToList();
    }

    [Benchmark]
    public async Task SpreadCheetah()
    {
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Book1");

        var cells = new DataCell[NumberOfColumns];
        var rows = _rows;

        foreach (var values in rows)
        {
            for (var i = 0; i < NumberOfColumns; ++i)
            {
                cells[i] = new DataCell(values[i]);
            }

            await spreadsheet.AddRowAsync(cells);
        }

        await spreadsheet.FinishAsync();
    }
}
