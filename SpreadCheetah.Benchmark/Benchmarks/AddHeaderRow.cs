using BenchmarkDotNet.Attributes;
using SpreadCheetah.Styling;

namespace SpreadCheetah.Benchmark.Benchmarks;

[MemoryDiagnoser]
public class AddHeaderRow
{
    private static readonly string[] HeaderNames =
    [
        "Id",
        "First name",
        "Last name",
        "Date of birth",
        "E-mail",
        "Role"
    ];

    [Params(100000)]
    public int NumberOfRows { get; set; }

    [Benchmark]
    public async Task SpreadCheetahAddHeaderRow()
    {
        var options = new SpreadCheetahOptions { DefaultDateTimeFormat = null };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, options);
        await spreadsheet.StartWorksheetAsync("Book1");
        var style = new Style { Font = { Bold = true } };
        var styleId = spreadsheet.AddStyle(style);

        for (var i = 0; i < NumberOfRows; ++i)
        {
            await spreadsheet.AddHeaderRowAsync(HeaderNames, styleId);
        }

        await spreadsheet.FinishAsync();
    }
}
