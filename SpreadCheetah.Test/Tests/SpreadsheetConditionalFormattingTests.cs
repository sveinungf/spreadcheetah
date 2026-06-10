using SpreadCheetah.ConditionalFormatting;
using SpreadCheetah.Helpers;
using SpreadCheetah.Styling;
using SpreadCheetah.Test.Extensions;
using SpreadCheetah.Test.Helpers;
using System.Drawing;
using System.IO.Compression;
using SpreadsheetAssert = SpreadCheetah.TestHelpers.SpreadsheetAssert;

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
    public async Task Spreadsheet_ConditionalFormatting_MultipleUniqueValuesRulesForSingleCell()
    {
        // Arrange
        const string cellReference = "B2";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var uniqueValues = ConditionalFormatRule.UniqueValues();

        List<UniqueValuesFormatRule> rules =
        [
            uniqueValues.WithStyle(new ConditionalFormatStyle { Fill = { Color = Color.Red } }),
            uniqueValues.WithStyle(new ConditionalFormatStyle { Font = { Bold = true } }),
            uniqueValues.WithStyle(new ConditionalFormatStyle { Format = "0.00" })
        ];

        // Act
        foreach (var rule in rules)
        {
            spreadsheet.AddConditionalFormatRule(cellReference, rule);
        }

        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(rules.Count, sheet.ConditionalFormatRules.Count);
        Assert.All(sheet.ConditionalFormatRules, x => Assert.True(x.IsUniqueValuesRule));
        Assert.All(sheet.ConditionalFormatRules, x => Assert.Equal(cellReference, x.CellRangeReference));
    }

    [Fact]
    public async Task Spreadsheet_ConditionalFormatting_MultipleUniqueValuesRulesForSingleCellHasExpectedSheetXml()
    {
        // Arrange
        const string cellReference = "B2";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var uniqueValues = ConditionalFormatRule.UniqueValues();

        List<UniqueValuesFormatRule> rules =
        [
            uniqueValues.WithStyle(new ConditionalFormatStyle { Fill = { Color = Color.Red } }),
            uniqueValues.WithStyle(new ConditionalFormatStyle { Font = { Bold = true } }),
            uniqueValues.WithStyle(new ConditionalFormatStyle { Format = "0.00" })
        ];

        // Act
        foreach (var rule in rules)
        {
            spreadsheet.AddConditionalFormatRule(cellReference, rule);
        }

        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var zip = await ZipArchive.CreateAsync(stream, Token);
        using var sheet1Xml = await zip.GetSheet1XmlStreamAsync(Token);
        await VerifyXml(sheet1Xml);
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
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        var style = new ConditionalFormatStyle
        {
            Fill = { Color = Color.FromArgb(10) },
            Format = "0.00",
            Font =
            {
                Bold = true,
                Color = Color.FromArgb(20),
                Italic = true,
                Strikethrough = true,
                Underline = ConditionalFormatUnderline.Single
            },
            Border =
            {
                Bottom = { BorderStyle = ConditionalFormatBorderStyle.Thin, Color = Color.FromArgb(30) },
                Left = { BorderStyle = ConditionalFormatBorderStyle.Dashed, Color = Color.FromArgb(40) },
                Right = { BorderStyle = ConditionalFormatBorderStyle.Dotted, Color = Color.FromArgb(50) },
                Top = { BorderStyle = ConditionalFormatBorderStyle.Hair, Color = Color.FromArgb(60) }
            }
        };

        // Act
        var rule = ConditionalFormatRule.UniqueValues().WithStyle(style);
        spreadsheet.AddConditionalFormatRule("A1", rule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRule = Assert.Single(sheet.ConditionalFormatRules);
        Assert.Equal(style.Fill.Color, actualRule.Style.Fill.Color);
        Assert.Equal(style.Format, actualRule.Style.NumberFormat.CustomFormat);
        Assert.Equal(style.Font.Color, actualRule.Style.Font.Color);
        Assert.Equal(Underline.Single, actualRule.Style.Font.Underline);
        Assert.True(actualRule.Style.Font.Bold);
        Assert.True(actualRule.Style.Font.Italic);
        Assert.True(actualRule.Style.Font.Strikethrough);
        Assert.Equal(BorderStyle.Thin, actualRule.Style.Border.BottomStyle);
        Assert.Equal(BorderStyle.Dashed, actualRule.Style.Border.LeftStyle);
        Assert.Equal(BorderStyle.Dotted, actualRule.Style.Border.RightStyle);
        Assert.Equal(BorderStyle.Hair, actualRule.Style.Border.TopStyle);
        Assert.Equal(style.Border.Bottom.Color, actualRule.Style.Border.BottomColor);
        Assert.Equal(style.Border.Left.Color, actualRule.Style.Border.LeftColor);
        Assert.Equal(style.Border.Right.Color, actualRule.Style.Border.RightColor);
        Assert.Equal(style.Border.Top.Color, actualRule.Style.Border.TopColor);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_ConditionalFormatting_RuleWithDefaultStyleHasExpectedSheetXml(bool explicitlyInitialized)
    {
        // Arrange
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new ConditionalFormatStyle();

        if (explicitlyInitialized)
        {
            style = new ConditionalFormatStyle
            {
                Border = new()
                {
                    Bottom = new(),
                    Left = new(),
                    Right = new(),
                    Top = new()
                },
                Fill = new(),
                Font = new()
            };
        }

        // Act
        var rule = ConditionalFormatRule.UniqueValues().WithStyle(style);
        spreadsheet.AddConditionalFormatRule("A1", rule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var zip = await ZipArchive.CreateAsync(stream, Token);
        using var sheet1Xml = await zip.GetSheet1XmlStreamAsync(Token);
        var verifySettings = new VerifySettings();
        verifySettings.IgnoreParametersForVerified();
        await VerifyXml(sheet1Xml, verifySettings);
    }

    [Fact]
    public async Task Spreadsheet_ConditionalFormatting_RuleWithDefaultStyleHasExpectedStylesXml()
    {
        // Arrange
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new ConditionalFormatStyle();

        // Act
        var rule = ConditionalFormatRule.UniqueValues().WithStyle(style);
        spreadsheet.AddConditionalFormatRule("A1", rule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var zip = await ZipArchive.CreateAsync(stream, Token);
        using var stylesXml = await zip.GetStylesXmlStreamAsync(Token);
        await VerifyXml(stylesXml);
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
    public async Task Spreadsheet_ConditionalFormatting_RuleWithBorderStyle(ConditionalFormatBorderStyle borderStyle)
    {
        // Arrange
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        var color = Color.FromArgb(255, 200, 0);
        var style = new ConditionalFormatStyle
        {
            Border = { Bottom = { BorderStyle = borderStyle, Color = color } }
        };

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
}
