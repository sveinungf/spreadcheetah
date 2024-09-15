# SpreadCheetah

[![Nuget](https://img.shields.io/nuget/v/SpreadCheetah?logo=nuget)](https://www.nuget.org/packages/SpreadCheetah)
[![GitHub Workflow Status (with event)](https://img.shields.io/github/actions/workflow/status/sveinungf/spreadcheetah/dotnet.yml?logo=github)](https://github.com/sveinungf/spreadcheetah/actions/workflows/dotnet.yml)
[![Codecov](https://img.shields.io/codecov/c/gh/sveinungf/spreadcheetah?logo=codecov)](https://app.codecov.io/gh/sveinungf/spreadcheetah)

SpreadCheetah is a high-performance .NET library for generating spreadsheet (Microsoft Excel XLSX) files.

## Features
- Performance (see benchmarks below)
- Low memory allocation (see benchmarks below)
- Async APIs
- No dependency to Microsoft Excel
- Targeting .NET Standard 2.0 for .NET Framework and earlier versions of .NET Core
- Targeting .NET 6 and newer for more optimizations
- Trimmable and NativeAOT compatible

SpreadCheetah is designed to create spreadsheet files in a forward-only manner.
That means worksheets from left to right, rows from top to bottom, and row cells from left to right.
This allows for creating spreadsheet files in a streaming manner, while also keeping a low memory footprint.

Most basic spreadsheet functionality is supported, such as cells with different data types, basic styling, and formulas. More advanced functionality will be added in future releases. See the list of currently supported spreadsheet functionality [in the wiki](https://github.com/sveinungf/spreadcheetah/wiki#supported-spreadsheet-functionality).

See [Releases](https://github.com/sveinungf/spreadcheetah/releases) for release notes.

## How to install
SpreadCheetah is available as a [package on nuget.org](https://www.nuget.org/packages/SpreadCheetah). The NuGet package targets .NET Standard 2.0 as well as newer versions of .NET. The .NET Standard 2.0 version is intended for backwards compatibility (.NET Framework and earlier versions of .NET Core). More optimizations are enabled when targeting newer versions of .NET.

## Basic usage
```cs
using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
{
    // A spreadsheet must contain at least one worksheet.
    await spreadsheet.StartWorksheetAsync("Sheet 1");

    // Cells are inserted row by row.
    var row = new List<Cell>
    {
        new Cell("Answer to the ultimate question:"),
        new Cell(42)
    };

    // Rows are inserted from top to bottom.
    await spreadsheet.AddRowAsync(row);

    // Remember to call Finish before disposing.
    // This is important to properly finalize the XLSX file.
    await spreadsheet.FinishAsync();
}
```

### Other examples
- [Writing to a file](https://github.com/sveinungf/spreadcheetah-samples/blob/main/SpreadCheetahSamples/WriteToFile.cs)
- [Styling basics](https://github.com/sveinungf/spreadcheetah-samples/blob/main/SpreadCheetahSamples/StylingBasics.cs)
- [Formula basics](https://github.com/sveinungf/spreadcheetah-samples/blob/main/SpreadCheetahSamples/FormulaBasics.cs)
- [DateTime and formatting](https://github.com/sveinungf/spreadcheetah-samples/blob/main/SpreadCheetahSamples/DateTimeAndFormatting.cs)
- [Data Validations](https://github.com/sveinungf/spreadcheetah-samples/blob/main/SpreadCheetahSamples/DataValidations.cs)
- [Adding an embedded image](https://github.com/sveinungf/spreadcheetah/wiki/Adding-an-embedded-image)
- [Performance tips](https://github.com/sveinungf/spreadcheetah-samples/blob/main/SpreadCheetahSamples/PerformanceTips.cs)

## Using the Source Generator
[Source Generators](https://devblogs.microsoft.com/dotnet/introducing-c-source-generators) is a feature in the C# compiler. SpreadCheetah includes a source generator that makes it easier to create rows from objects. It is used in a similar way to the [`System.Text.Json` source generator](https://devblogs.microsoft.com/dotnet/try-the-new-system-text-json-source-generator/):
```cs
namespace MyNamespace;

// A plain old C# class which we want to add as a row in a worksheet.
// The source generator will pick the properties with public getters.
// The order of the properties will decide the order of the cells.
public class MyObject
{
    public string Question { get; set; }
    public int Answer { get; set; }
}
```

The source generator will be instructed to generate code by defining a partial class like this:
```cs
using SpreadCheetah.SourceGeneration;

namespace MyNamespace;

[WorksheetRow(typeof(MyObject))]
public partial class MyObjectRowContext : WorksheetRowContext;
```

During build, the type will be analyzed and an implementation of the context class will be created. We can then create a row from an object by calling `AddAsRowAsync` with the object and the context type as parameters:
```cs
await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
await spreadsheet.StartWorksheetAsync("Sheet 1");

var myObj = new MyObject
{
    Question = "How many Rings of Power were there?",
    Answer = 20
};

await spreadsheet.AddAsRowAsync(myObj, MyObjectRowContext.Default.MyObject);

await spreadsheet.FinishAsync();
```

Here is a peek at part of the code that was generated for this example. As you can see, the generated code also takes advantage of using a pooled array to avoid memory allocations:
```cs
// <auto-generated />
private static async ValueTask AddAsRowInternalAsync(Spreadsheet spreadsheet, MyObject obj, CancellationToken token)
{
    var cells = ArrayPool<DataCell>.Shared.Rent(2);
    try
    {
        cells[0] = new DataCell(obj.Question);
        cells[1] = new DataCell(obj.Answer);
        await spreadsheet.AddRowAsync(cells.AsMemory(0, 2), token).ConfigureAwait(false);
    }
    finally
    {
        ArrayPool<DataCell>.Shared.Return(cells, true);
    }
}
```

The source generator can generate rows from classes, records, and structs. It can be used in all supported .NET versions, including .NET Framework, however the C# version must be 8.0 or greater.
More features of the source generator can be seen in the [wiki](https://github.com/sveinungf/spreadcheetah/wiki/Source-generator).

## Benchmarks
The benchmark results here have been collected using [BenchmarkDotNet](https://github.com/dotnet/benchmarkdotnet) with the following system configuration:

``` ini
BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.4894/22H2/2022Update)
Intel Core i5-8600K CPU 3.60GHz (Coffee Lake), 1 CPU, 6 logical and 6 physical cores
.NET SDK 9.0.100-preview.7.24407.12
  [Host]             : .NET 8.0.8 (8.0.824.36612), X64 RyuJIT AVX2
  .NET 6.0           : .NET 6.0.33 (6.0.3324.36610), X64 RyuJIT AVX2
  .NET 8.0           : .NET 8.0.8 (8.0.824.36612), X64 RyuJIT AVX2
  .NET Framework 4.8 : .NET Framework 4.8.1 (4.8.9261.0), X64 RyuJIT VectorSize=256

InvocationCount=1  UnrollFactor=1
```

These libraries have been used in the comparison benchmarks:
| Library                                                | Version |
|--------------------------------------------------------|--------:|
| SpreadCheetah                                          |  1.18.0 |
| [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) |  2.20.0 |
| [ClosedXML](https://github.com/ClosedXML/ClosedXML)    | 0.102.3 |
| [EPPlusFree](https://github.com/rimland/EPPlus)        | 4.5.3.8 |

> Disclaimer: The libraries have different feature sets compared to each other.
> SpreadCheetah can only create spreadsheets, while the other libraries used in this comparison
> can also open spreadsheets. SpreadCheetah is also a newer library and has been designed from
> the ground up to utilize many of the newer performance related features in .NET. The other
> libraries have longer history and need to take backwards compatibility into account.
> Keep this in mind when evaluating the results.

The code executed in the benchmark creates a worksheet of 20 000 rows and 10 columns filled
with string values. Some of these libraries have multiple ways of achieving the same result,
but to make this a fair comparison the idea is to use the most efficient approach for each library.
The code is available [in the Benchmark project](https://github.com/sveinungf/spreadcheetah/blob/main/SpreadCheetah.Benchmark/Benchmarks/StringCells.cs).

<details open>
<summary><h3>.NET 8</h3></summary>

|                    Library |         Mean |     Allocated |
|----------------------------|-------------:|--------------:|
|          **SpreadCheetah** | **21.22 ms** |   **6.33 KB** |
|    Open XML (SAX approach) |    185.66 ms |  66 037.32 KB |
|                 EPPlusFree |    358.31 ms | 195 610.91 KB |
|    Open XML (DOM approach) |    701.17 ms | 182 916.73 KB |
|                  ClosedXML |    739.16 ms | 529 203.20 KB |
</details>


<details>
<summary><h3>.NET 6</h3></summary>

|                    Library |         Mean |     Allocated |
|----------------------------|-------------:|--------------:|
|          **SpreadCheetah** | **25.05 ms** |   **6.52 KB** |
|    Open XML (SAX approach) |    234.12 ms |  66 052.24 KB |
|                 EPPlusFree |    407.38 ms | 195 791.84 KB |
|    Open XML (DOM approach) |    803.30 ms | 182 926.09 KB |
|                  ClosedXML |    874.41 ms | 529 844.80 KB |
</details>


<details>
<summary><h3>.NET Framework 4.8</h3></summary>

|                    Library |         Mean |     Allocated |
|----------------------------|-------------:|--------------:|
|          **SpreadCheetah** | **72.34 ms** | **152.23 KB** |
|    Open XML (SAX approach) |    408.03 ms |  43 317.24 KB |
|                 EPPlusFree |    622.80 ms | 286 141.77 KB |
|    Open XML (DOM approach) |  1,070.28 ms | 161 067.34 KB |
|                  ClosedXML |  1,319.16 ms | 509 205.80 KB |
</details>
