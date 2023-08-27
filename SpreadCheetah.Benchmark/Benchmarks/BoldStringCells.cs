using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using SpreadCheetah.Styling;

namespace SpreadCheetah.Benchmark.Benchmarks;

[SimpleJob(RuntimeMoniker.Net70)]
[MemoryDiagnoser]
public class BoldStringCells
{
    private const int NumberOfColumns = 10;
    private const string SheetName = "Book1";

    private List<List<string>> _values = null!;

    [Params(20000)]
    public int NumberOfRows { get; set; }

    [GlobalSetup]
    public void IterationSetup()
    {
        _values = new List<List<string>>(NumberOfRows);

        for (var row = 1; row <= NumberOfRows; ++row)
        {
            var columns = Enumerable.Range(1, NumberOfColumns).ToList();
            _values.Add(columns.ConvertAll(x => $"{x}-{row}"));
        }
    }

    [Benchmark]
    public async Task SpreadCheetah()
    {
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync(SheetName);

        var style = new Style();
        style.Font.Bold = true;
        var styleId = spreadsheet.AddStyle(style);
        var cells = new StyledCell[NumberOfColumns];

        for (var row = 0; row < NumberOfRows; ++row)
        {
            for (var col = 0; col < NumberOfColumns; ++col)
            {
                var value = _values[row][col];
                cells[col] = new StyledCell(value, styleId);
            }

            await spreadsheet.AddRowAsync(cells);
        }

        await spreadsheet.FinishAsync();
    }
}
