using SpreadCheetah.SourceGenerator.Test.Models.CellFormat;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Assertions;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class CellFormatTests
{
    [Fact]
    public async Task CellFormat_ClassWithStandardFormat()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var obj = new ClassWithCellStandardFormat { Price = 199.90m };

        // Act
        await spreadsheet.AddAsRowAsync(obj, CellFormatContext.Default.ClassWithCellStandardFormat);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(obj.Price, sheet["A1"].DecimalValue);
        Assert.Equal(StandardNumberFormat.TwoDecimalPlaces, sheet["A1"].Style.NumberFormat.StandardFormat);
    }

    [Fact]
    public async Task CellFormat_ClassWithCustomFormat()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var obj = new ClassWithCellCustomFormat { Price = 199.90m };

        // Act
        await spreadsheet.AddAsRowAsync(obj, CellFormatContext.Default.ClassWithCellCustomFormat);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(obj.Price, sheet["A1"].DecimalValue);
        Assert.Equal("#.0#", sheet["A1"].Style.NumberFormat.CustomFormat);
    }

    [Theory, CombinatorialData]
    public async Task CellFormat_ClassWithDateTimeFormat(bool withDefaultDateTimeFormat)
    {
        // Arrange
        using var stream = new MemoryStream();
        var options = withDefaultDateTimeFormat
            ? new SpreadCheetahOptions()
            : new SpreadCheetahOptions { DefaultDateTimeFormat = null };

        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var obj = new ClassWithDateTimeCellFormat
        {
            FromDate = DateTime.UtcNow.AddDays(-1),
            ToDate = DateTime.UtcNow.AddDays(1)
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj, CellFormatContext.Default.ClassWithDateTimeCellFormat);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(withDefaultDateTimeFormat, sheet["A1"].Style.NumberFormat.CustomFormat is not null);
        Assert.Equal(StandardNumberFormat.DateAndTime, sheet["B1"].Style.NumberFormat.StandardFormat);
    }

    [Fact]
    public async Task CellFormat_ClassWithMultipleCellFormats()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var obj = new ClassWithMultipleCellFormats
        {
            FromDate = DateTime.UtcNow,
            Id = 6889,
            Price = 199.90m
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj, CellFormatContext.Default.ClassWithMultipleCellFormats);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal("#.00", sheet["A1"].Style.NumberFormat.CustomFormat);
        Assert.Equal(StandardNumberFormat.LongDate, sheet["B1"].Style.NumberFormat.StandardFormat);
        Assert.Equal("#.00", sheet["C1"].Style.NumberFormat.CustomFormat);
    }
}
