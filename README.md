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
BenchmarkDotNet=v0.12.1, OS=Windows 10.0.19041.572 (2004/?/20H1)
Intel Core i5-8600K CPU 3.60GHz (Coffee Lake), 1 CPU, 6 logical and 6 physical cores
.NET Core SDK=5.0.100-preview.8.20417.9
  [Host]        : .NET Core 3.1.9 (CoreCLR 4.700.20.47201, CoreFX 4.700.20.47203), X64 RyuJIT
  .NET 4.8      : .NET Framework 4.8 (4.8.4250.0), X64 RyuJIT
  .NET Core 3.1 : .NET Core 3.1.9 (CoreCLR 4.700.20.47201, CoreFX 4.700.20.47203), X64 RyuJIT
```

The code executed in the benchmark is filling a worksheet with 20000 rows and 10 columns. I've also implemented the same use case in other spreadsheet libraries for comparison. You can also see the how they perform in .NET Framework and .NET Core.

|       Library |       Runtime |        Mean |     Error |    StdDev |    Allocated |
|-------------- |-------------- |------------:|----------:|----------:|-------------:|
| **SpreadCheetah** |      **.NET 4.8** | **61.93 ms** | **0.257 ms** | **0.215 ms** | **152.23 KB** |
|       EpPlus4 |      .NET 4.8 |   711.77 ms |  1.779 ms |  1.577 ms | 331002.49 KB |
|    OpenXmlSax |      .NET 4.8 |   384.09 ms |  0.986 ms |  0.922 ms |  10407.93 KB |
|    OpenXmlDom |      .NET 4.8 |   655.18 ms |  1.173 ms |  1.097 ms |  53369.68 KB |
|     ClosedXml |      .NET 4.8 | 2,209.31 ms |  7.445 ms |  6.964 ms | 557844.45 KB |
| **SpreadCheetah** | **.NET Core 3.1** | **29.42 ms** | **0.215 ms** | **0.201 ms** |  **51.93 KB** |
|       EpPlus4 | .NET Core 3.1 |   579.76 ms |  0.659 ms |  0.550 ms | 245572.94 KB |
|    OpenXmlSax | .NET Core 3.1 |   201.78 ms |  0.687 ms |  0.536 ms |  33235.98 KB |
|    OpenXmlDom | .NET Core 3.1 |   546.74 ms |  2.881 ms |  2.554 ms |  76051.36 KB |
|     ClosedXml | .NET Core 3.1 | 1,964.78 ms | 35.960 ms | 33.637 ms |  531324.8 KB |
