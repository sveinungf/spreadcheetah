using SpreadCheetah.SourceGenerator.Test.Models.CellStyle;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Assertions;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class CellStyleTests
{
    [Fact]
    public async Task CellStyle_ClassWithSingleAttribute()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var style = new Style { Font = { Bold = true } };
        spreadsheet.AddStyle(style, "Price style");
        var obj = new ClassWithCellStyle { Price = 199.90m };

        // Act
        await spreadsheet.AddAsRowAsync(obj, CellStyleContext.Default.ClassWithCellStyle);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(obj.Price, sheet["A1"].DecimalValue);
        Assert.True(sheet["A1"].Style.Font.Bold);
    }

    [Fact]
    public async Task CellStyle_ClassWithMultipleAttributes()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var boldStyle = new Style { Font = { Bold = true } };
        spreadsheet.AddStyle(boldStyle, "Name");
        var italicStyle = new Style { Font = { Italic = true } };
        spreadsheet.AddStyle(italicStyle, "Age");
        var obj = new ClassWithMultipleCellStyles
        {
            FirstName = "Ola",
            MiddleName = "Jonsen",
            LastName = "Nordmann",
            Age = 42
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj, CellStyleContext.Default.ClassWithMultipleCellStyles);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.True(sheet["A1"].Style.Font.Bold);
        Assert.False(sheet["A1"].Style.Font.Italic);
        Assert.False(sheet["B1"].Style.Font.Bold);
        Assert.False(sheet["B1"].Style.Font.Italic);
        Assert.True(sheet["C1"].Style.Font.Bold);
        Assert.False(sheet["C1"].Style.Font.Italic);
        Assert.False(sheet["D1"].Style.Font.Bold);
        Assert.True(sheet["D1"].Style.Font.Italic);
    }
}