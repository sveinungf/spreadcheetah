using Polyfills;
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
    public async Task Spreadsheet_ConditionalFormatting_DuplicateValuesRule()
    {
        // Arrange
        const string cellReference = "A1:A10";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var fillColor = Color.FromArgb(255, 255, 0, 0);
        var style = new ConditionalFormatStyle { Fill = { Color = fillColor } };

        // Act
        var rule = ConditionalFormatRule.DuplicateValues().WithStyle(style);
        spreadsheet.AddConditionalFormatRule(cellReference, rule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRule = Assert.Single(sheet.ConditionalFormatRules);
        Assert.True(actualRule.IsDuplicateValuesRule);
        Assert.Equal(cellReference, actualRule.CellRangeReference);
        Assert.Equal(fillColor, actualRule.Style.Fill.Color);
    }

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
    public async Task Spreadsheet_ConditionalFormatting_DuplicateAndUniqueValuesRulesForSameCellRange()
    {
        // Arrange
        const string cellReference = "A1:A10";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style1 = new ConditionalFormatStyle { Font = { Bold = true } };
        var style2 = new ConditionalFormatStyle { Font = { Italic = true } };

        // Act
        var rule1 = ConditionalFormatRule.DuplicateValues().WithStyle(style1);
        var rule2 = ConditionalFormatRule.UniqueValues().WithStyle(style2);
        spreadsheet.AddConditionalFormatRule(cellReference, rule1);
        spreadsheet.AddConditionalFormatRule(cellReference, rule2);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRules = sheet.ConditionalFormatRules;
        var actualDuplicateValuesRule = Assert.Single(actualRules, x => x.IsDuplicateValuesRule);
        Assert.Equal(style1.Font.Bold, actualDuplicateValuesRule.Style.Font.Bold);
        var actualUniqueValuesRule = Assert.Single(actualRules, x => x.IsUniqueValuesRule);
        Assert.Equal(style2.Font.Italic, actualUniqueValuesRule.Style.Font.Italic);
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
    public async Task Spreadsheet_ConditionalFormatting_ManyDuplicateValuesRules()
    {
        // Arrange
        const int count = SpreadsheetConstants.MaxNumberOfConditionalFormatRules;
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new ConditionalFormatStyle { Fill = { Color = Color.Red } };

        // Act
        var rule = ConditionalFormatRule.DuplicateValues().WithStyle(style);

        for (var i = 0; i < count; i++)
        {
            var cellReference = $"A{i + 1}:B{i + 1}";
            spreadsheet.AddConditionalFormatRule(cellReference, rule);
        }

        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(count, sheet.ConditionalFormatRules.Count);
        Assert.All(sheet.ConditionalFormatRules, x => Assert.True(x.IsDuplicateValuesRule));
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
    public async Task Spreadsheet_ConditionalFormatting_MultipleUniqueValuesRulesForMultipleCellsHaveExpectedSheetXml()
    {
        // Arrange
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
        foreach (var (index, rule) in rules.Index())
        {
            var cellReference = $"C{index + 1}";
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
    public async Task Spreadsheet_ConditionalFormatting_TooManyRules()
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
    public async Task Spreadsheet_ConditionalFormatting_MultipleRulesWithFontUnderlineStyle(ConditionalFormatUnderline underline)
    {
        // Arrange
        const int count = 100;
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        var style = new ConditionalFormatStyle { Font = { Underline = underline } };

        // Act
        var rule = ConditionalFormatRule.UniqueValues().WithStyle(style);
        for (var i = 0; i < count; i++)
        {
            spreadsheet.AddConditionalFormatRule($"A{i + 1}:B{i + 1}", rule);
        }

        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var expected = (Underline)underline;
        Assert.All(sheet.ConditionalFormatRules, x => Assert.Equal(expected, x.Style.Font.Underline));
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

    [Fact]
    public async Task Spreadsheet_ConditionalFormatting_MatchesFormulaA1Rule()
    {
        // Arrange
        const string cellReference = "B2:C3";
        const string formula = "A1>5";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new ConditionalFormatStyle { Fill = { Color = Color.Red } };

        // Act
        var rule = ConditionalFormatRule.MatchesFormula(new Formula(formula)).WithStyle(style);
        spreadsheet.AddConditionalFormatRule(cellReference, rule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRule = Assert.Single(sheet.ConditionalFormatRules);
        Assert.True(actualRule.IsMatchesFormulaRule);
        Assert.Equal(formula, actualRule.Formula);
    }

    [Theory]
    [InlineData("B2", "RC[-1]", "A2")]
    [InlineData("C3", "RC[1]", "D3")]
    [InlineData("B2", "RC[0]", "B2")]
    [InlineData("A1:A10", "RC[1]", "B1")]
    [InlineData("B2:C3", "R1C1", "$A$1")]
    [InlineData("B2:C3", "R[1]C[1]", "C3")]
    [InlineData("B2", "SUM(RC[-1]:RC[1])>10", "SUM(A2:C2)>10")]
    [InlineData("D4", "R[10]C[5]", "I14")]
    [InlineData("B2", "SUM(R2:R4)", "SUM($2:$4)")]
    [InlineData("B2", "C[-1]&RC[0]", "A:A&B2")]
    public async Task Spreadsheet_ConditionalFormatting_MatchesFormulaR1C1Rule(
        string cellReference, string r1c1Formula, string expectedA1Formula)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new ConditionalFormatStyle { Fill = { Color = Color.Red } };

        // Act
        var rule = ConditionalFormatRule.MatchesFormula(Formula.R1C1(r1c1Formula)).WithStyle(style);
        spreadsheet.AddConditionalFormatRule(cellReference, rule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRule = Assert.Single(sheet.ConditionalFormatRules);
        Assert.True(actualRule.IsMatchesFormulaRule);
        Assert.Equal(expectedA1Formula, actualRule.Formula);
    }

    [Fact]
    public async Task Spreadsheet_ConditionalFormatting_MatchesFormulaRuleForSingleCell()
    {
        // Arrange
        const string cellReference = "B2";
        const string formula = "A1>5";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var fillColor = Color.FromArgb(255, 255, 0, 0);
        var style = new ConditionalFormatStyle { Fill = { Color = fillColor } };

        // Act
        var rule = ConditionalFormatRule.MatchesFormula(new Formula(formula)).WithStyle(style);
        spreadsheet.AddConditionalFormatRule(cellReference, rule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRule = Assert.Single(sheet.ConditionalFormatRules);
        Assert.True(actualRule.IsMatchesFormulaRule);
        Assert.Equal(cellReference, actualRule.CellRangeReference);
        Assert.Equal(formula, actualRule.Formula);
        Assert.Equal(fillColor, actualRule.Style.Fill.Color);
    }

    [Fact]
    public async Task Spreadsheet_ConditionalFormatting_MultipleMatchesFormulaRulesForSingleCell()
    {
        // Arrange
        const string cellReference = "B2";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        List<MatchesFormulaFormatRule> rules =
        [
            ConditionalFormatRule.MatchesFormula(new Formula("A1>5")).WithStyle(new() { Fill = { Color = Color.Red } }),
            ConditionalFormatRule.MatchesFormula(new Formula("A1<10")).WithStyle(new() { Font = { Bold = true } }),
            ConditionalFormatRule.MatchesFormula(Formula.R1C1("RC[-1]")).WithStyle(new() { Format = "0.00" })
        ];

        // Act
        foreach (var rule in rules)
        {
            spreadsheet.AddConditionalFormatRule(cellReference, rule);
        }

        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRules = sheet.ConditionalFormatRules;
        Assert.Equal(rules.Count, actualRules.Count);
        Assert.All(actualRules, x => Assert.True(x.IsMatchesFormulaRule));
        Assert.All(actualRules, x => Assert.Equal(cellReference, x.CellRangeReference));
        Assert.Equal("A1>5", actualRules[0].Formula);
        Assert.Equal("A1<10", actualRules[1].Formula);
        Assert.Equal("A2", actualRules[2].Formula);
        Assert.Equal(Color.FromArgb(255, 255, 0, 0), actualRules[0].Style.Fill.Color);
        Assert.True(actualRules[1].Style.Font.Bold);
        Assert.Equal("0.00", actualRules[2].Style.NumberFormat.CustomFormat);
    }

    [Fact]
    public async Task Spreadsheet_ConditionalFormatting_MatchesFormulaAndUniqueValuesRulesForSameCellRange()
    {
        // Arrange
        const string cellReference = "A1:A10";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style1 = new ConditionalFormatStyle { Fill = { Color = Color.FromArgb(255, 255, 0, 0) } };
        var style2 = new ConditionalFormatStyle { Font = { Bold = true } };

        // Act
        var formula = new Formula("A1>5");
        var formulaRule = ConditionalFormatRule.MatchesFormula(formula).WithStyle(style1);
        var uniqueValuesRule = ConditionalFormatRule.UniqueValues().WithStyle(style2);
        spreadsheet.AddConditionalFormatRule(cellReference, formulaRule);
        spreadsheet.AddConditionalFormatRule(cellReference, uniqueValuesRule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRules = sheet.ConditionalFormatRules;
        Assert.All(actualRules, x => Assert.Equal(cellReference, x.CellRangeReference));
        var actualFormulaRule = Assert.Single(actualRules, x => x.IsMatchesFormulaRule);
        Assert.Equal("A1>5", actualFormulaRule.Formula);
        Assert.Equal(style1.Fill.Color, actualFormulaRule.Style.Fill.Color);
        var actualUniqueValuesRule = Assert.Single(actualRules, x => x.IsUniqueValuesRule);
        Assert.Equal(style2.Font.Bold, actualUniqueValuesRule.Style.Font.Bold);
    }

    [Fact]
    public async Task Spreadsheet_ConditionalFormatting_ManyMatchesFormulaRules()
    {
        // Arrange
        const int count = SpreadsheetConstants.MaxNumberOfConditionalFormatRules;
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new ConditionalFormatStyle { Fill = { Color = Color.Red } };

        // Act
        var formula = new Formula("A1>5");
        var rule = ConditionalFormatRule.MatchesFormula(formula).WithStyle(style);

        for (var i = 0; i < count; i++)
        {
            var cellReference = $"A{i + 1}:B{i + 1}";
            spreadsheet.AddConditionalFormatRule(cellReference, rule);
        }

        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(count, sheet.ConditionalFormatRules.Count);
        Assert.All(sheet.ConditionalFormatRules, x => Assert.True(x.IsMatchesFormulaRule));
    }

    [Theory]
    [InlineData("A1<5")]
    [InlineData("AND(A1>2,A2<10)")]
    [InlineData("A1&B1")]
    [InlineData("A1=\"text\"")]
    [InlineData("A1<>\"don't\"")]
    public async Task Spreadsheet_ConditionalFormatting_MatchesFormulaRuleWithXmlSpecialCharacters(string formula)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new ConditionalFormatStyle { Fill = { Color = Color.Red } };

        // Act
        var rule = ConditionalFormatRule.MatchesFormula(new Formula(formula)).WithStyle(style);
        spreadsheet.AddConditionalFormatRule("A1", rule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRule = Assert.Single(sheet.ConditionalFormatRules);
        Assert.True(actualRule.IsMatchesFormulaRule);
        Assert.Equal(formula, actualRule.Formula);
    }

    [Fact]
    public async Task Spreadsheet_ConditionalFormatting_MatchesFormulaRuleWithLongFormula()
    {
        // Arrange
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new ConditionalFormatStyle { Fill = { Color = Color.Red } };
        var formula = "A1>" + new string('1', 10000);

        // Act
        var rule = ConditionalFormatRule.MatchesFormula(new Formula(formula)).WithStyle(style);
        spreadsheet.AddConditionalFormatRule("A1", rule);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualRule = Assert.Single(sheet.ConditionalFormatRules);
        Assert.True(actualRule.IsMatchesFormulaRule);
        Assert.Equal(formula, actualRule.Formula);
    }

    [Fact]
    public async Task Spreadsheet_ConditionalFormatting_MatchesFormulaR1C1RuleWithOutOfRangeReference()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new ConditionalFormatStyle { Fill = { Color = Color.Red } };

        // Act
        var formula = Formula.R1C1("RC[-1]");
        var rule = ConditionalFormatRule.MatchesFormula(formula).WithStyle(style);
        spreadsheet.AddConditionalFormatRule("A1", rule);

        // Assert
        var exception = await Assert.ThrowsAsync<SpreadCheetahException>(() => spreadsheet.FinishAsync(Token).AsTask());
        Assert.Equal(
            "The R1C1 formula 'RC[-1]' contains a reference that is outside the bounds of the worksheet when anchored to this cell.",
            exception.Message);
    }

    [Fact]
    public void Spreadsheet_ConditionalFormatting_MatchesFormulaRuleWithNullStyle()
    {
        // Arrange
        var rule = ConditionalFormatRule.MatchesFormula(new Formula("A1>5"));
        ConditionalFormatStyle style = null!;

        // Act & Assert
        Assert.Throws<ArgumentNullException>(() => rule.WithStyle(style));
    }

    [Fact]
    public async Task Spreadsheet_ConditionalFormatting_MatchesFormulaRulesHaveExpectedSheetXml()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var formula1 = new Formula("A1>5");
        var formula2 = Formula.R1C1("RC[-1]");
        var rule1 = ConditionalFormatRule.MatchesFormula(formula1).WithStyle(new() { Fill = { Color = Color.Red } });
        var rule2 = ConditionalFormatRule.MatchesFormula(formula2).WithStyle(new() { Font = { Bold = true } });

        // Act
        spreadsheet.AddConditionalFormatRule("D6", rule1);
        spreadsheet.AddConditionalFormatRule("B2:C3", rule2);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var zip = await ZipArchive.CreateAsync(stream, Token);
        using var sheet1Xml = await zip.GetSheet1XmlStreamAsync(Token);
        await VerifyXml(sheet1Xml);
    }
}
