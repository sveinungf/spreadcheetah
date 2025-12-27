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
                    K = random.NextBoolean() ? baseDateTime.AddSeconds(random.Next()) : null,
                    L = (long)random.Next() * random.Next(10, 1000000000)
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

        foreach (var row in _rows)
        {
            cells[0] = new DataCell(row.A);
            cells[1] = new DataCell(row.B);
            cells[2] = new DataCell(row.C);
            cells[3] = new DataCell(row.D);
            cells[4] = new DataCell(row.E);
            cells[5] = new DataCell(row.F);
            cells[6] = new DataCell(row.G);
            cells[7] = new DataCell(row.H);
            cells[8] = new DataCell(row.I);
            cells[9] = new DataCell(row.J);
            cells[10] = new DataCell(row.K);

            await spreadsheet.AddRowAsync(cells);
        }

        await spreadsheet.FinishAsync();
    }
}
