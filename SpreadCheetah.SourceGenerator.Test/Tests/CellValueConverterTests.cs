using SpreadCheetah.SourceGenerator.Test.Models.CellValueConverter;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Assertions;
using System.Globalization;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class CellValueConverterTests
{
    [Fact]
    public async Task CellValueConverter_ClassWithReusedConverter()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var obj = new ClassWithReusedConverter
        {
            FirstName = "Ola",
            MiddleName = null,
            LastName = "Nordmann",
            Gpa = 3.1m
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj, CellValueConverterContext.Default.ClassWithReusedConverter);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal("OLA", sheet["A1"].StringValue);
        Assert.Equal("NORDMANN", sheet["C1"].StringValue);
    }

    [Fact]
    public async Task CellValueConverter_ClassWithGenericConverter()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var obj = new ClassWithGenericConverter
        {
            FirstName = "Ola",
            MiddleName = null,
            LastName = "Nordmann",
            Gpa = 3.1m
        };

        // Act
        CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
        await spreadsheet.AddAsRowAsync(obj, CellValueConverterContext.Default.ClassWithGenericConverter);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal("-", sheet["B1"].StringValue);
        Assert.Equal("3.1", sheet["D1"].StringValue);
    }

    [Fact]
    public async Task CellValueConverter_TwoClassesUsingTheSameConverter()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var obj1 = new ClassWithCellValueConverter { Name = "John" };
        var obj2 = new ClassWithReusedConverter
        {
            FirstName = "Ola",
            MiddleName = null,
            LastName = "Nordmann",
            Gpa = 3.1m
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj1, CellValueConverterContext.Default.ClassWithCellValueConverter);
        await spreadsheet.AddAsRowAsync(obj2, CellValueConverterContext.Default.ClassWithReusedConverter);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal("JOHN", sheet["A1"].StringValue);
        Assert.Equal("OLA", sheet["A2"].StringValue);
        Assert.Equal("NORDMANN", sheet["C2"].StringValue);
    }

    [Fact]
    public async Task CellValueConverter_ClassWithConverterAndCellStyle()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var style = new Style { Font = { Bold = true } };
        spreadsheet.AddStyle(style, "ID style");
        var obj = new ClassWithCellValueConverterAndCellStyle { Id = "Abc123" };

        // Act
        await spreadsheet.AddAsRowAsync(obj, CellValueConverterContext.Default.ClassWithCellValueConverterAndCellStyle);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal("ABC123", sheet["A1"].StringValue);
        Assert.True(sheet["A1"].Style.Font.Bold);
    }
    
    [Fact]
    public async Task CellValueConverter_ClassWithConverterOnCustomType()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var style = new Style { Font = { Bold = true } };
        spreadsheet.AddStyle(style, "PercentType");
        var obj = new ClassWithCellValueConverterOnCustomType
        {
            Property = "Abc123", ComplexProperty = null, PercentType = new Percent(123)
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj, CellValueConverterContext.Default.ClassWithCellValueConverterOnCustomType);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal("Abc123", sheet["A1"].StringValue);
        
        Assert.Equal("-", sheet["B1"].StringValue);
        
        Assert.Equal(123, sheet["C1"].DecimalValue);
        Assert.True(sheet["C1"].Style.Font.Bold);
    }
}
