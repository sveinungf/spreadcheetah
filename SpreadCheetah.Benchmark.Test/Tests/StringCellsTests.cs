using SpreadCheetah.Benchmark.Benchmarks;
using SpreadCheetah.TestHelpers.Assertions;
using Xunit;

namespace SpreadCheetah.Benchmark.Test.Tests;

public sealed class StringCellsTests : IDisposable
{
    private const int NumberOfColumns = 10;
    private const int NumberOfRows = 20000;
    private readonly StringCells _stringCells;
    private readonly ITestOutputHelper _output;

    public StringCellsTests(ITestOutputHelper output)
    {
        _output = output;
        _stringCells = new StringCells
        {
            NumberOfColumns = NumberOfColumns,
            NumberOfRows = NumberOfRows,
            Stream = new MemoryStream()
        };

        _stringCells.GlobalSetup();
    }

    public void Dispose()
    {
        _stringCells.Dispose();
    }

    [Fact]
    public async Task StringCells_SpreadCheetah_CorrectCellValues()
    {
        // Act
        await _stringCells.SpreadCheetah();

        // Assert
        WriteOutput();
        AssertCellValuesEqual();
    }

    [Fact]
    public void StringCells_EpPlus4_CorrectCellValues()
    {
        // Act
        _stringCells.EpPlus4();

        // Assert
        WriteOutput();
        AssertCellValuesEqual();
    }

    [Fact]
    public void StringCells_OpenXmlSax_CorrectCellValues()
    {
        // Act
        _stringCells.OpenXmlSax();

        // Assert
        WriteOutput();
        AssertCellValuesEqual();
    }

    [Fact]
    public void StringCells_OpenXmlDom_CorrectCellValues()
    {
        // Act
        _stringCells.OpenXmlDom();

        // Assert
        WriteOutput();
        AssertCellValuesEqual();
    }

    [Fact]
    public void StringCells_ClosedXml_CorrectCellValues()
    {
        // Act
        _stringCells.ClosedXml();

        // Assert
        WriteOutput();
        AssertCellValuesEqual();
    }

    private void WriteOutput()
    {
        _output.WriteLine("Stream length: " + _stringCells.Stream.Length);
    }

    private void AssertCellValuesEqual()
    {
        using var sheet = SpreadsheetAssert.SingleSheet(_stringCells.Stream);

        var values = _stringCells.Values;
        Assert.Equal(values.Count, sheet.RowCount);

        foreach (var (r, rowValues) in values.Index())
        {
            var row = sheet.Row(r + 1).Cells.ToList();
            Assert.Equal(rowValues.Count, row.Count);

            foreach (var (c, cellValue) in rowValues.Index())
            {
                Assert.Equal(cellValue, row[c].StringValue);
            }
        }
    }
}
