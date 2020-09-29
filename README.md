# SpreadCheetah
SpreadCheetah is a high-performance .NET library for generating spreadsheet (Microsoft Excel XLSX) files.


## Features
- Performance (see benchmarks below)
- Low memory allocation (see benchmarks below)
- Async APIs
- No dependency to Microsoft Excel
- Targeting .NET Standard 2.0 (for .NET Framework 4.6.1 and later)
- Free and open source!

SpreadCheetah is designed to create spreadsheet files in a forward-only manner.
That means worksheets from left to right, rows from top to bottom, and row cells from left to right.
This allows for creating spreadsheet files in a streaming manner, while also keeping a low memory footprint.


## How to install
SpreadCheetah is available as a [NuGet package](https://www.nuget.org/packages/SpreadCheetah). The NuGet package targets both .NET Standard 2.0 and .NET Standard 2.1.
The .NET Standard 2.0 version is just intended for backwards compatibility. Use the .NET Standard 2.1 version to achieve the best performance.


## Usage
```cs
using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
{
    // A spreadsheet must contain at least one worksheet.
    await spreadsheet.PutNextWorksheetAsync("Sheet 1");

    // Cells are inserted row by row.
    var row = new List<Cell>();
    row.Add(new Cell("Answer to the ultimate question:"));
    row.Add(new Cell(42));

    // Rows are inserted from top to bottom.
    await spreadsheet.AddRowAsync(row);

    // Remember to call Finish before disposing. This is important to properly finalize the XLSX file.
    await spreadsheet.FinishAsync();
}
```

## Benchmarks
The benchmark results here have been collected using [Benchmark.NET](https://github.com/dotnet/benchmarkdotnet) with the following system configuration:

``` ini
BenchmarkDotNet=v0.12.1, OS=Windows 10.0.19041.508 (2004/?/20H1)
Intel Core i5-8600K CPU 3.60GHz (Coffee Lake), 1 CPU, 6 logical and 6 physical cores
.NET Core SDK=5.0.100-rc.1.20452.10
  [Host]        : .NET Core 3.1.8 (CoreCLR 4.700.20.41105, CoreFX 4.700.20.41903), X64 RyuJIT
  .NET 4.8      : .NET Framework 4.8 (4.8.4220.0), X64 RyuJIT
  .NET Core 3.1 : .NET Core 3.1.8 (CoreCLR 4.700.20.41105, CoreFX 4.700.20.41903), X64 RyuJIT
```

The code executed in the benchmark is filling a worksheet with 20000 rows and 10 columns. I've also implemented the same use case in other spreadsheet libraries for comparison. You can also see the how they perform in .NET Framework and .NET Core.

|       Library |       Runtime |        Mean |     Error |    StdDev |    Allocated |
|-------------- |-------------- |------------:|----------:|----------:|-------------:|
| SpreadCheetah |      .NET 4.8 |    64.80 ms |  1.133 ms |  1.059 ms |    152.23 KB |
|       EpPlus4 |      .NET 4.8 |   964.72 ms |  4.769 ms |  4.461 ms | 331005.53 KB |
|    OpenXmlSax |      .NET 4.8 |   356.21 ms |  1.807 ms |  1.691 ms |  10407.93 KB |
|    OpenXmlDom |      .NET 4.8 |   654.49 ms |  1.687 ms |  1.578 ms |  53369.68 KB |
|     ClosedXml |      .NET 4.8 | 2,195.07 ms |  6.942 ms |  5.420 ms | 557807.62 KB |
| SpreadCheetah | .NET Core 3.1 |    30.36 ms |  0.519 ms |  0.485 ms |     48.44 KB |
|       EpPlus4 | .NET Core 3.1 |   841.52 ms |  5.861 ms |  5.482 ms | 245513.84 KB |
|    OpenXmlSax | .NET Core 3.1 |   265.28 ms |  4.838 ms |  4.525 ms |  33235.98 KB |
|    OpenXmlDom | .NET Core 3.1 |   581.86 ms |  3.618 ms |  3.385 ms |  76051.36 KB |
|     ClosedXml | .NET Core 3.1 | 1,966.03 ms | 28.659 ms | 26.807 ms | 531324.38 KB |
