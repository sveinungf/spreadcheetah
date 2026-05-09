using SpreadCheetah.ConditionalFormatting;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Assertions;
using System.Drawing;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetConditionalFormattingTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public async Task Spreadsheet_ConditionalFormatting_UniqueValuesRule()
    {
        // Arrange
        const string cellReference = "A1:A10";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var fillColor = Color.FromArgb(255, 255, 0, 0);
        var style = new Style { Fill = { Color = fillColor } };

        // Act
        var rule = ConditionalFormatRule.UniqueValues().WithStyle(style);
        spreadsheet.AddConditionalFormatRule(cellReference, rule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRule = Assert.Single(sheet.ConditionalFormatRules);
        Assert.True(actualRule.IsUniqueValuesRule);
        Assert.Equal(cellReference, actualRule.CellRangeReference);
        Assert.Equal(fillColor, actualRule.Style.Fill.Color);
    }

    [Fact]
    public async Task Spreadsheet_ConditionalFormatting_UniqueValuesRuleForSingleCell()
    {
        // Arrange
        const string cellReference = "A1";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var fillColor = Color.FromArgb(255, 255, 0, 0);
        var style = new Style { Fill = { Color = fillColor } };

        // Act
        var rule = ConditionalFormatRule.UniqueValues().WithStyle(style);
        spreadsheet.AddConditionalFormatRule(cellReference, rule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRule = Assert.Single(sheet.ConditionalFormatRules);
        Assert.True(actualRule.IsUniqueValuesRule);
        Assert.Equal(cellReference, actualRule.CellRangeReference);
        Assert.Equal(fillColor, actualRule.Style.Fill.Color);
    }

    [Fact]
    public async Task Spreadsheet_ConditionalFormatting_ManyUniqueValuesRules()
    {
        // Arrange
        const int count = SpreadsheetConstants.MaxNumberOfConditionalFormatRules;
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new Style { Fill = { Color = Color.Red } };

        // Act
        var rule = ConditionalFormatRule.UniqueValues().WithStyle(style);

        for (var i = 0; i < count; i++)
        {
            var cellReference = $"A{i + 1}:B{i + 1}";
            spreadsheet.AddConditionalFormatRule(cellReference, rule);
        }

        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(count, sheet.ConditionalFormatRules.Count);
        Assert.All(sheet.ConditionalFormatRules, x => Assert.True(x.IsUniqueValuesRule));
    }

    [Fact]
    public async Task Spreadsheet_ConditionalFormatting_TooManyUniqueValuesRules()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new Style { Fill = { Color = Color.Green } };
        var rule = ConditionalFormatRule.UniqueValues().WithStyle(style);

        for (var i = 0; i < SpreadsheetConstants.MaxNumberOfConditionalFormatRules; i++)
        {
            var cellReference = $"A{i + 1}:B{i + 1}";
            spreadsheet.AddConditionalFormatRule(cellReference, rule);
        }

        // Act
        var exception = Assert.Throws<SpreadCheetahException>(() => spreadsheet.AddConditionalFormatRule("C1:D2", rule));
        Assert.Equal("Can't add more than 16384 conditional format rules to a worksheet.", exception.Message);
    }

    [Fact]
    public async Task Spreadsheet_ConditionalFormatting_RuleWithAllStyleElements()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        var fillColor = Color.FromArgb(255, 255, 200, 0);
        const string numberFormat = "0.00";
        // TODO: Add more style elements.
        var style = new Style
        {
            Fill = { Color = fillColor },
            Format = NumberFormat.Custom(numberFormat)
        };

        // Act
        var rule = ConditionalFormatRule.UniqueValues().WithStyle(style);
        spreadsheet.AddConditionalFormatRule("A1", rule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRule = Assert.Single(sheet.ConditionalFormatRules);
        Assert.Equal(fillColor, actualRule.Style.Fill.Color);
        Assert.Equal(numberFormat, actualRule.Style.NumberFormat.CustomFormat);
    }

    [Fact]
    public async Task Spreadsheet_ConditionalFormatting_RuleWithLongNumberFormat()
    {
        // Arrange
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        var numberFormat = new string('"', 255);
        var style = new Style { Format = NumberFormat.Custom(numberFormat) };

        // Act
        var rule = ConditionalFormatRule.UniqueValues().WithStyle(style);
        spreadsheet.AddConditionalFormatRule("A1", rule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRule = Assert.Single(sheet.ConditionalFormatRules);
        Assert.Equal(numberFormat, actualRule.Style.NumberFormat.CustomFormat);
    }
}
