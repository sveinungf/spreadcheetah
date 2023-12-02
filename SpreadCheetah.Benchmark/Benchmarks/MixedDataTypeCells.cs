using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using SpreadCheetah.Benchmark.Helpers;

namespace SpreadCheetah.Benchmark.Benchmarks;

[SimpleJob(RuntimeMoniker.Net48)]
[SimpleJob(RuntimeMoniker.Net70)]
[SimpleJob(RuntimeMoniker.Net80)]
[MemoryDiagnoser]
public class MixedDataTypeCells
{
    private RowItem[] _rows = null!;

    [Params(500000)]
    public int NumberOfRows { get; set; }

    [GlobalSetup]
    public void GlobalSetup()
    {
        _rows = Enumerable.Range(0, NumberOfRows)
            .Select(row =>
            {
                var even = row % 2 == 0;
                return new RowItem(
                    row,
                    "Ola-" + row,
                    row + "-Nordmann",
                    1950 + row / 1000,
                    5.67 + row % 10,
                    even,
                    even ? "Norway" : "Sweden",
                    !even,
                    0.991f + row / 10000.0,
                    even ? -23 : 23);
            }).ToArray();
    }

    [Benchmark]
    public async Task SpreadCheetah()
    {
        var options = new SpreadCheetahOptions { DefaultDateTimeFormat = null };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, options);
        await spreadsheet.StartWorksheetAsync("Book1");

        var cells = new DataCell[10];
        var rows = _rows;

        for (var row = 0; row < rows.Length; ++row)
        {
            var rowItem = rows[row];
            cells[0] = new DataCell(rowItem.A);
            cells[1] = new DataCell(rowItem.B);
            cells[2] = new DataCell(rowItem.C);
            cells[3] = new DataCell(rowItem.D);
            cells[4] = new DataCell(rowItem.E);
            cells[5] = new DataCell(rowItem.F);
            cells[6] = new DataCell(rowItem.G);
            cells[7] = new DataCell(rowItem.H);
            cells[8] = new DataCell(rowItem.I);
            cells[9] = new DataCell(rowItem.J);

            await spreadsheet.AddRowAsync(cells);
        }

        await spreadsheet.FinishAsync();
    }
}
