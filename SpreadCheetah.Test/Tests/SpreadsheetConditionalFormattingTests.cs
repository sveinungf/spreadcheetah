using SpreadCheetah.ConditionalFormatting;
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
}
