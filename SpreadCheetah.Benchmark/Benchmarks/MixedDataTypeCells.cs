using BenchmarkDotNet.Attributes;
using SpreadCheetah.Benchmark.Helpers;

namespace SpreadCheetah.Benchmark.Benchmarks;

[MemoryDiagnoser]
public class MixedDataTypeCells
{
    private RowItem[] _rows = null!;

    [Params(500000)]
    public int NumberOfRows { get; set; }

    [GlobalSetup]
    public void GlobalSetup()
    {
        var random = new Random(42);
        var baseDateTime = new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Utc);

        _rows =
        [
            .. Enumerable.Range(0, NumberOfRows)
            .Select(row =>
            {
                var even = row % 2 == 0;
                return new RowItem
                {
                    A = row,
                    B = "Ola-" + row,
                    C = random.NextBoolean() ? row + "-Nordmann" : null,
                    D = random.NextBoolean() ? null : 1950 + row / 1000,
                    E = 5.67 + row % 10,
                    F = even,
                    G = even ? "Norway" : "Sweden",
                    H = !even,
                    I = random.NextBoolean() ? 0.991f + row / 10000.0 : null,
                    J = even ? -23 : 23,
                    K = random.NextBoolean() ? baseDateTime.AddSeconds(random.Next()) : null
                };
            })
        ];
    }

    [Benchmark]
    public async Task SpreadCheetah()
    {
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Book1");

        var cells = new DataCell[11];
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
            cells[10] = new DataCell(rowItem.K);

            await spreadsheet.AddRowAsync(cells);
        }

        await spreadsheet.FinishAsync();
    }
}
