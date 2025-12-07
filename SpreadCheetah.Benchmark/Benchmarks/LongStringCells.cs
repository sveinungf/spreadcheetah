using BenchmarkDotNet.Attributes;
using Bogus;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadCheetah.Benchmark.Benchmarks;

[MemoryDiagnoser]
public class LongStringCells
{
    private static readonly SpreadCheetahOptions Options = new()
    {
        BufferSize = SpreadCheetahOptions.MinimumBufferSize
    };

    private List<List<string>> _rows = [];

    [Params(20)]
    public int NumberOfColumns { get; set; }

    [Params(1000)]
    public int NumberOfRows { get; set; }

    [GlobalSetup]
    public void GlobalSetup()
    {
        var randomizer = new Randomizer(localSeed: 42);

        _rows = Enumerable
            .Range(0, NumberOfRows)
            .Select(_ => Enumerable
                .Range(0, NumberOfColumns)
                .Select(_ => randomizer.Words(1000))
                .ToList())
            .ToList();
    }

    [Benchmark]
    public async Task SpreadCheetah()
    {
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, Options);
        await spreadsheet.StartWorksheetAsync("Book1");

        var cells = new DataCell[NumberOfColumns];

        foreach (var row in _rows)
        {
            for (var col = 0; col < NumberOfColumns; col++)
            {
                cells[col] = new DataCell(row[col]);
            }

            await spreadsheet.AddRowAsync(cells);
        }

        await spreadsheet.FinishAsync();
    }
}
