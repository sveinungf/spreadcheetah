# SpreadCheetah
SpreadCheetah is a high-performance .NET library for generating spreadsheet (Microsoft Excel XLSX) files.
Highlights:
- Performance (benchmarks will be added soon)
- Low memory allocation (benchmarks will be added soon)
- Async APIs
- No dependency to Microsoft Excel
- Targets .NET Standard 2.0
- Free and open source!

SpreadCheetah is designed to create spreadsheet files in a forward-only manner.
That means worksheets from left to right, rows from top to bottom, and row cells from left to right.
This allows for creating spreadsheet files in a streaming manner, while also keeping a low memory footprint.

## How to install
SpreadCheetah is available as a NuGet package. The NuGet package targets both .NET Standard 2.0 and .NET Standard 2.1.
The .NET Standard 2.0 version is just intended for backwards compatibility. Use the .NET Standard 2.1 version to achieve the best performance.
