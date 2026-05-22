using SpreadCheetah.ConditionalFormatting;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers;
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
        var style = new ConditionalFormatStyle { Fill = { Color = fillColor } };

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
        var style = new ConditionalFormatStyle { Fill = { Color = fillColor } };

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
        var style = new ConditionalFormatStyle { Fill = { Color = Color.Red } };

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
        var style = new ConditionalFormatStyle { Fill = { Color = Color.Green } };
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
        var fontColor = Color.FromArgb(255, 10, 0, 0);
        // TODO: Add more style elements.
        var style = new ConditionalFormatStyle
        {
            Fill = { Color = fillColor },
            Format = "0.00",
            Font =
            {
                Bold = true,
                Color = fontColor,
                Italic = true,
                Strikethrough = true,
                Underline = ConditionalFormatUnderline.Single
            }
        };

        // Act
        var rule = ConditionalFormatRule.UniqueValues().WithStyle(style);
        spreadsheet.AddConditionalFormatRule("A1", rule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRule = Assert.Single(sheet.ConditionalFormatRules);
        Assert.Equal(fillColor, actualRule.Style.Fill.Color);
        Assert.Equal(style.Format, actualRule.Style.NumberFormat.CustomFormat);
        Assert.Equal(fontColor, actualRule.Style.Font.Color);
        Assert.Equal(Underline.Single, actualRule.Style.Font.Underline);
        Assert.True(actualRule.Style.Font.Bold);
        Assert.True(actualRule.Style.Font.Italic);
        Assert.True(actualRule.Style.Font.Strikethrough);
    }

    [Fact]
    public async Task Spreadsheet_ConditionalFormatting_RuleWithLongNumberFormatStyle()
    {
        // Arrange
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        var style = new ConditionalFormatStyle { Format = new string('"', 255) };

        // Act
        var rule = ConditionalFormatRule.UniqueValues().WithStyle(style);
        spreadsheet.AddConditionalFormatRule("A1", rule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRule = Assert.Single(sheet.ConditionalFormatRules);
        Assert.Equal(style.Format, actualRule.Style.NumberFormat.CustomFormat);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_ConditionalFormatting_RuleWithFontUnderlineStyle(ConditionalFormatUnderline underline)
    {
        // Arrange
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        var style = new ConditionalFormatStyle { Font = { Underline = underline } };

        // Act
        var rule = ConditionalFormatRule.UniqueValues().WithStyle(style);
        spreadsheet.AddConditionalFormatRule("A1:A5", rule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRule = Assert.Single(sheet.ConditionalFormatRules);
        Assert.Equal((Underline)underline, actualRule.Style.Font.Underline);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_ConditionalFormatting_RuleWithBottomBorderStyle(ConditionalFormatBorderStyle borderStyle)
    {
        // Arrange
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        var color = Color.FromArgb(255, 255, 200, 0);
        var style = new ConditionalFormatStyle { Border = { Bottom = { BorderStyle = borderStyle, Color = color } } };

        // Act
        var rule = ConditionalFormatRule.UniqueValues().WithStyle(style);
        spreadsheet.AddConditionalFormatRule("A1:A5", rule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRule = Assert.Single(sheet.ConditionalFormatRules);
        Assert.Equal((BorderStyle)borderStyle, actualRule.Style.Border.BottomStyle);

        if (borderStyle != ConditionalFormatBorderStyle.None)
            Assert.Equal(color, actualRule.Style.Border.BottomColor);
    }

    // TODO: More tests for borders.
}
