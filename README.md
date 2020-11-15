# SpreadCheetah

[![Nuget](https://img.shields.io/nuget/v/SpreadCheetah)](https://www.nuget.org/packages/SpreadCheetah)

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
    await spreadsheet.StartWorksheetAsync("Sheet 1");

    // Cells are inserted row by row.
    var row = new List<DataCell>();
    row.Add(new DataCell("Answer to the ultimate question:"));
    row.Add(new DataCell(42));

    // Rows are inserted from top to bottom.
    await spreadsheet.AddRowAsync(row);

    // Remember to call Finish before disposing. This is important to properly finalize the XLSX file.
    await spreadsheet.FinishAsync();
}
```

## Benchmarks
The benchmark results here have been collected using [Benchmark.NET](https://github.com/dotnet/benchmarkdotnet) with the following system configuration:

``` ini
BenchmarkDotNet=v0.12.1, OS=Windows 10.0.19041.630 (2004/?/20H1)
Intel Core i5-8600K CPU 3.60GHz (Coffee Lake), 1 CPU, 6 logical and 6 physical cores
.NET Core SDK=5.0.100
  [Host]        : .NET Core 5.0.0 (CoreCLR 5.0.20.51904, CoreFX 5.0.20.51904), X64 RyuJIT
  .NET 4.8      : .NET Framework 4.8 (4.8.4250.0), X64 RyuJIT
  .NET Core 3.1 : .NET Core 3.1.9 (CoreCLR 4.700.20.47201, CoreFX 4.700.20.47203), X64 RyuJIT
  .NET Core 5.0 : .NET Core 5.0.0 (CoreCLR 5.0.20.51904, CoreFX 5.0.20.51904), X64 RyuJIT
```

The code executed in the benchmark is filling a worksheet with 20000 rows and 10 columns. I've also implemented the same use case in other spreadsheet libraries for comparison. You can also see the how they perform in .NET Framework and .NET Core.

|       Library |       Runtime |        Mean |     Error |    StdDev |     Allocated |
|-------------- |-------------- |------------:|----------:|----------:|--------------:|
| **SpreadCheetah** |      **.NET 4.8** |    **64.33 ms** |  **0.484 ms** |  **0.404 ms** |     **136.23 KB** |
|       EpPlus4 |      .NET 4.8 |   571.33 ms |  6.479 ms |  6.061 ms |  286368.81 KB |
|    OpenXmlSax |      .NET 4.8 |   330.48 ms |  2.532 ms |  2.368 ms |   35485.21 KB |
|    OpenXmlDom |      .NET 4.8 |   794.46 ms |  9.945 ms |  9.302 ms |  135947.15 KB |
|     ClosedXml |      .NET 4.8 | 2,158.61 ms | 18.147 ms | 16.975 ms |  558507.04 KB |
| **SpreadCheetah** | **.NET Core 3.1** |    **32.69 ms** |  **0.295 ms** |  **0.262 ms** |      **68.95 KB** |
|       EpPlus4 | .NET Core 3.1 |   453.88 ms |  8.675 ms | 10.972 ms |  216141.84 KB |
|    OpenXmlSax | .NET Core 3.1 |   195.84 ms |  2.985 ms |  2.792 ms |   58245.29 KB |
|    OpenXmlDom | .NET Core 3.1 |   644.66 ms |  2.634 ms |  2.335 ms |  158085.77 KB |
|     ClosedXml | .NET Core 3.1 | 1,897.42 ms | 13.222 ms | 11.721 ms |  531326.55 KB |
