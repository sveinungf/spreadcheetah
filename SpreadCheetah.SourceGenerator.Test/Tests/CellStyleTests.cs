using System.Reflection;
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Helpers;
using SpreadCheetah.SourceGenerator.Test.Models.CellStyle;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Assertions;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class CellStyleTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public async Task CellStyle_ClassWithSingleAttribute()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new Style { Font = { Bold = true } };
        spreadsheet.AddStyle(style, "Price style");
        var obj = new ClassWithCellStyle { Price = 199.90m };

        // Act
        await spreadsheet.AddAsRowAsync(obj, CellStyleContext.Default.ClassWithCellStyle, Token);
        await spreadsheet.FinishAsync(Token);

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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
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
        await spreadsheet.AddAsRowAsync(obj, CellStyleContext.Default.ClassWithMultipleCellStyles, Token);
        await spreadsheet.FinishAsync(Token);

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

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task CellStyle_ClassWithAttributeOnDateTimeProperty(bool withDefaultDateTimeFormat)
    {
        // Arrange
        using var stream = new MemoryStream();
        var options = withDefaultDateTimeFormat
            ? new SpreadCheetahOptions()
            : new SpreadCheetahOptions { DefaultDateTimeFormat = null };

        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new Style { Font = { Bold = true } };
        spreadsheet.AddStyle(style, "Created style");
        var obj = new ClassWithCellStyleOnDateTimeProperty { CreatedDate = new DateTime(2022, 10, 23, 15, 16, 17, DateTimeKind.Utc) };

        // Act
        await spreadsheet.AddAsRowAsync(obj, CellStyleContext.Default.ClassWithCellStyleOnDateTimeProperty, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(obj.CreatedDate, sheet["A1"].DateTimeValue);
        var actualStyle = sheet["A1"].Style;
        Assert.True(actualStyle.Font.Bold);
        Assert.Equal(withDefaultDateTimeFormat, actualStyle.NumberFormat.CustomFormat is not null);
    }

    [Fact]
    public async Task CellStyle_ClassWithAttributeOnTruncatedProperty()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new Style { Font = { Bold = true } };
        spreadsheet.AddStyle(style, "Name style");
        var obj = new ClassWithCellStyleOnTruncatedProperty { Name = "Ola Nordmann" };

        // Act
        await spreadsheet.AddAsRowAsync(obj, CellStyleContext.Default.ClassWithCellStyleOnTruncatedProperty, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(obj.Name[..10], sheet["A1"].StringValue);
        Assert.True(sheet["A1"].Style.Font.Bold);
    }

    [Fact]
    public async Task CellStyle_MultipleTypesWithAttributeInSameWorksheet()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var nameStyle = new Style { Font = { Italic = true } };
        var priceStyle = new Style { Font = { Bold = true } };
        spreadsheet.AddStyle(nameStyle, "Name style");
        spreadsheet.AddStyle(priceStyle, "Price style");
        var obj1 = new ClassWithCellStyle { Price = 199.90m };
        var obj2 = new RecordWithCellStyle { Name = "Tom" };

        // Act
        await spreadsheet.AddAsRowAsync(obj1, CellStyleContext.Default.ClassWithCellStyle, Token);
        await spreadsheet.AddAsRowAsync(obj2, CellStyleContext.Default.RecordWithCellStyle, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(obj1.Price, sheet["A1"].DecimalValue);
        Assert.Equal(obj2.Name, sheet["A2"].StringValue);
        Assert.True(sheet["A1"].Style.Font.Bold);
        Assert.True(sheet["A2"].Style.Font.Italic);
    }

    [Fact]
    public async Task CellStyle_MultipleTypesWithAttributeInMultipleWorksheets()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);

        // Act
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        var priceStyle = new Style { Font = { Bold = true } };
        spreadsheet.AddStyle(priceStyle, "Price style");
        var obj1 = new ClassWithCellStyle { Price = 199.90m };
        await spreadsheet.AddAsRowAsync(obj1, CellStyleContext.Default.ClassWithCellStyle, Token);

        await spreadsheet.StartWorksheetAsync("Sheet 2", token: Token);
        var nameStyle = new Style { Font = { Italic = true } };
        spreadsheet.AddStyle(nameStyle, "Name style");
        var obj2 = new RecordWithCellStyle { Name = "Tom" };
        await spreadsheet.AddAsRowAsync(obj2, CellStyleContext.Default.RecordWithCellStyle, Token);

        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheets = SpreadsheetAssert.Sheets(stream);
        Assert.Equal(2, sheets.Count);

        var sheet1 = sheets[0];
        Assert.Equal(obj1.Price, sheet1["A1"].DecimalValue);
        Assert.True(sheet1["A1"].Style.Font.Bold);

        var sheet2 = sheets[1];
        Assert.Equal(obj2.Name, sheet2["A1"].StringValue);
        Assert.True(sheet2["A1"].Style.Font.Italic);
    }

    [Fact]
    public async Task CellStyle_ClassWithMissingStyleName()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var obj = new ClassWithCellStyle { Price = 199.90m };

        // Act
        var exception = await Record.ExceptionAsync(async () => await spreadsheet.AddAsRowAsync(obj, CellStyleContext.Default.ClassWithCellStyle, Token));

        // Assert
        var actual = Assert.IsType<SpreadCheetahException>(exception);
        Assert.Contains(nameof(Spreadsheet.AddStyle), actual.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task CellStyle_WorksheetRowDependencyInfoCreatedOnlyOnce()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new Style { Font = { Bold = true } };
        spreadsheet.AddStyle(style, "Price style");
        var typeInfo = CellStyleContext.Default.ClassWithCellStyle;
        var initialDependencyInfo = spreadsheet.GetOrCreateWorksheetRowDependencyInfo(typeInfo);

        // Act
        await spreadsheet.AddAsRowAsync(new ClassWithCellStyle { Price = 10m }, typeInfo, Token);
        await spreadsheet.AddAsRowAsync(new ClassWithCellStyle { Price = 20m }, typeInfo, Token);
        await spreadsheet.AddAsRowAsync(new ClassWithCellStyle { Price = 30m }, typeInfo, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        var subsequentDependencyInfo = spreadsheet.GetOrCreateWorksheetRowDependencyInfo(typeInfo);
        Assert.Same(initialDependencyInfo, subsequentDependencyInfo);
    }

    [Fact]
    public void CellStyle_ClassWithCellStyle_CanReadStyleName()
    {
        // Arrange
        var properties = typeof(ClassWithCellStyle).ToPropertyDictionary();
        var property = properties[nameof(ClassWithCellStyle.Price)];

        // Act
        var attribute = property.GetCustomAttribute<CellStyleAttribute>();

        // Assert
        Assert.NotNull(attribute);
        Assert.Equal("Price style", attribute.StyleName);
    }
}