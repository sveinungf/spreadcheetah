using ClosedXML.Excel;
using SpreadCheetah.Styling;
using SpreadCheetah.Test.Helpers;
using SpreadCheetah.Test.Helpers.Backporting;
using System.Drawing;
using Xunit;
using Font = SpreadCheetah.Styling.Font;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetStyledRowTests
{
    [Theory]
    [MemberData(nameof(TestData.StyledCellAndValueTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_StyledCellWithValue(CellValueType valueType, bool isNull, CellType cellType)
    {
        // Arrange
        object? value;
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Font.Bold = true;
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(cellType, valueType, isNull, styleId, out value);

            // Act
            await spreadsheet.AddRowAsync(styledCell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(value?.ToString() ?? "", actualCell.Value.ToString());
        Assert.True(actualCell.Style.Font.Bold);
    }

    public static IEnumerable<object?[]> TrueAndFalse => TestData.CombineWithStyledCellTypes(true, false);

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_BoldCellWithStringValue(bool bold, CellType type)
    {
        // Arrange
        const string cellValue = "Bold test";
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Font.Bold = bold;
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(cellValue, actualCell.Value);
        Assert.Equal(bold, actualCell.Style.Font.Bold);
    }

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_SameBoldStyleCells(bool bold, CellType type)
    {
        // Arrange
        const string firstCellValue = "First";
        const string secondCellValue = "Second";
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Font.Bold = bold;
            var styleId = spreadsheet.AddStyle(style);

            var firstCell = CellFactory.Create(type, firstCellValue, styleId);
            var secondCell = CellFactory.Create(type, secondCellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(new[] { firstCell, secondCell });
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualFirstCell = worksheet.Cell(1, 1);
        var actualSecondCell = worksheet.Cell(1, 2);
        Assert.Equal(firstCellValue, actualFirstCell.Value);
        Assert.Equal(secondCellValue, actualSecondCell.Value);
        Assert.Equal(bold, actualFirstCell.Style.Font.Bold);
        Assert.Equal(bold, actualSecondCell.Style.Font.Bold);
    }

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_MixedBoldStyleCells(bool firstCellBold, CellType type)
    {
        // Arrange
        const string firstCellValue = "First";
        const string secondCellValue = "Second";
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Font.Bold = true;
            var styleId = spreadsheet.AddStyle(style);

            var firstCell = CellFactory.Create(type, firstCellValue, firstCellBold ? styleId : null);
            var secondCell = CellFactory.Create(type, secondCellValue, firstCellBold ? null : styleId);

            // Act
            await spreadsheet.AddRowAsync(new[] { firstCell, secondCell });
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualFirstCell = worksheet.Cell(1, 1);
        var actualSecondCell = worksheet.Cell(1, 2);
        Assert.Equal(firstCellValue, actualFirstCell.Value);
        Assert.Equal(secondCellValue, actualSecondCell.Value);
        Assert.Equal(firstCellBold, actualFirstCell.Style.Font.Bold);
        Assert.Equal(!firstCellBold, actualSecondCell.Style.Font.Bold);
    }

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_ItalicCellWithStringValue(bool italic, CellType type)
    {
        // Arrange
        const string cellValue = "Italic test";
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Font.Italic = italic;
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(cellValue, actualCell.Value);
        Assert.Equal(italic, actualCell.Style.Font.Italic);
    }

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_StrikethroughCellWithStringValue(bool strikethrough, CellType type)
    {
        // Arrange
        const string cellValue = "Italic test";
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Font.Strikethrough = strikethrough;
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(cellValue, actualCell.Value);
        Assert.Equal(strikethrough, actualCell.Style.Font.Strikethrough);
    }

    public static IEnumerable<object?[]> FontSizes() => TestData.CombineWithStyledCellTypes(
        8,
        11,
        11.5,
        72);

    [Theory]
    [MemberData(nameof(FontSizes))]
    public async Task Spreadsheet_AddRow_FontSizeCellWithStringValue(double size, CellType type)
    {
        // Arrange
        const string cellValue = "Font size test";
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Font.Size = size;
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(cellValue, actualCell.Value);
        Assert.Equal(size, actualCell.Style.Font.FontSize);
    }

    [Theory]
    [MemberData(nameof(TestData.StyledCellTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_FontColorCellWithStringValue(CellType type)
    {
        // Arrange
        const string cellValue = "Color test";
        var color = Color.Blue;
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Font.Color = color;
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(cellValue, actualCell.Value);
        Assert.Equal(color.ToArgb(), actualCell.Style.Font.FontColor.Color.ToArgb());
    }

    [Theory]
    [MemberData(nameof(TestData.StyledCellTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_FillColorCellWithStringValue(CellType type)
    {
        // Arrange
        const string cellValue = "Color test";
        var color = Color.Brown;
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Fill.Color = color;
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(cellValue, actualCell.Value);
        Assert.Equal(color.ToArgb(), actualCell.Style.Fill.BackgroundColor.Color.ToArgb());
    }

    public static IEnumerable<object?[]> FontNames() => TestData.CombineWithStyledCellTypes(
        "Arial",
        "SimSun-ExtB",
        "Times New Roman",
        null);

    [Theory]
    [MemberData(nameof(FontNames))]
    public async Task Spreadsheet_AddRow_FontNameCellWithStringValue(string? fontName, CellType type)
    {
        // Arrange
        const string cellValue = "Font name test";
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Font.Name = fontName;
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(cellValue, actualCell.Value);
        Assert.Equal(fontName ?? "Calibri", actualCell.Style.Font.FontName);
    }

    public static IEnumerable<object?[]> PredefinedNumberFormats() => TestData.CombineWithStyledCellTypes(
        NumberFormats.Fraction,
        NumberFormats.FractionTwoDenominatorPlaces,
        NumberFormats.General,
        NumberFormats.NoDecimalPlaces,
        NumberFormats.Percent,
        NumberFormats.PercentTwoDecimalPlaces,
        NumberFormats.Scientific,
        NumberFormats.ThousandsSeparator,
        NumberFormats.ThousandsSeparatorTwoDecimalPlaces,
        NumberFormats.TwoDecimalPlaces,
        null);

    [Theory]
    [MemberData(nameof(PredefinedNumberFormats))]
    public async Task Spreadsheet_AddRow_PredefinedNumberFormatCellWithStringValue(string? numberFormat, CellType type)
    {
        // Arrange
        const string cellValue = "Number format test";
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style { NumberFormat = numberFormat };
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell);
            await spreadsheet.FinishAsync();
        }

        var expectedNumberFormatId = NumberFormats.GetPredefinedNumberFormatId(numberFormat) ?? 0;

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(cellValue, actualCell.Value);
        Assert.Equal(expectedNumberFormatId, actualCell.Style.NumberFormat.NumberFormatId);
    }

    public static IEnumerable<object?[]> CustomNumberFormats() => TestData.CombineWithStyledCellTypes(
        "0.0000",
        @"0.0\ %",
        "[<=9999]0000;General",
        @"[<=99999999]##_ ##_ ##_ ##;\(\+##\)_ ##_ ##_ ##_ ##",
        @"_-* #,##0.0_-;\-* #,##0.0_-;_-* ""-""??_-;_-@_-");

    [Theory]
    [MemberData(nameof(CustomNumberFormats))]
    public async Task Spreadsheet_AddRow_CustomNumberFormatCellWithStringValue(string numberFormat, CellType type)
    {
        // Arrange
        const string cellValue = "Number format test";
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style { NumberFormat = numberFormat };
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(cellValue, actualCell.Value);
        Assert.Equal(numberFormat, actualCell.Style.NumberFormat.Format);
    }

    [Theory]
    [MemberData(nameof(TestData.StyledCellTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_MultipleStylesWithTheSameCustomNumberFormat(CellType type)
    {
        // Arrange
        const string cellValue = "Number format test";
        const string format = "0.0000";
        var styles = new Style[]
        {
            new() { NumberFormat = format },
            new() { NumberFormat = format, Fill = new Fill { Color = Color.Coral } },
            new() { NumberFormat = format, Font = new Font { Bold = true } }
        };

        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var cells = styles.Select(x => CellFactory.Create(type, cellValue, spreadsheet.AddStyle(x))).ToList();

            // Act
            await spreadsheet.AddRowAsync(cells);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCells = worksheet.CellsUsed();
        Assert.All(actualCells, x => Assert.Equal(cellValue, x.Value));
        Assert.All(actualCells, x => Assert.Equal(format, x.Style.NumberFormat.Format));
        Assert.Equal(styles.Length, actualCells.Count());
    }

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_DateTimeNumberFormat(bool withExplicitNumberFormat, CellType type)
    {
        // Arrange
        const string explicitNumberFormat = "yyyy";
        var expectedNumberFormat = withExplicitNumberFormat ? explicitNumberFormat : NumberFormats.DateTimeSortable;
        var value = new DateTime(2021, 2, 3, 4, 5, 6);
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Font.Italic = true;
            if (withExplicitNumberFormat)
                style.NumberFormat = explicitNumberFormat;

            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, value, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(value, actualCell.GetDateTime());
        Assert.Equal(expectedNumberFormat, actualCell.Style.NumberFormat.Format);
        Assert.True(actualCell.Style.Font.Italic);
    }

    public static IEnumerable<object?[]> BorderStyles() => TestData.CombineWithStyledCellTypes(EnumHelper.GetValues<BorderStyle>());

    [Theory]
    [MemberData(nameof(BorderStyles))]
    public async Task Spreadsheet_AddRow_BorderStyle(BorderStyle borderStyle, CellType type)
    {
        // Arrange
        const string cellValue = "Border style test";
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style { Border = new Border { Left = new EdgeBorder { BorderStyle = borderStyle } } };
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell);
            await spreadsheet.FinishAsync();
        }

        var expectedBorderStyle = borderStyle.GetClosedXmlBorderStyle();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        var actualBorder = actualCell.Style.Border;
        Assert.Equal(cellValue, actualCell.Value);
        Assert.Equal(expectedBorderStyle, actualBorder.LeftBorder);
        Assert.Equal(XLColor.Black, actualBorder.LeftBorderColor);
        Assert.Equal(XLBorderStyleValues.None, actualBorder.RightBorder);
        Assert.Equal(XLBorderStyleValues.None, actualBorder.TopBorder);
        Assert.Equal(XLBorderStyleValues.None, actualBorder.BottomBorder);
        Assert.Equal(XLBorderStyleValues.None, actualBorder.DiagonalBorder);
    }

    public static IEnumerable<object?[]> DiagonalBorderTypes() => TestData.CombineWithStyledCellTypes(EnumHelper.GetValues<DiagonalBorderType>());

    [Theory]
    [MemberData(nameof(DiagonalBorderTypes))]
    public async Task Spreadsheet_AddRow_DiagonalBorder(DiagonalBorderType borderType, CellType type)
    {
        // Arrange
        const string cellValue = "Border style test";
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style
            {
                Border = new Border
                {
                    Diagonal = new DiagonalBorder
                    {
                        Color = Color.Green,
                        BorderStyle = BorderStyle.Thin,
                        Type = borderType
                    }
                }
            };
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        var actualBorder = actualCell.Style.Border;
        Assert.Equal(cellValue, actualCell.Value);
        Assert.Equal(XLBorderStyleValues.Thin, actualBorder.DiagonalBorder);
        Assert.Equal(borderType.HasFlag(DiagonalBorderType.DiagonalUp), actualBorder.DiagonalUp);
        Assert.Equal(borderType.HasFlag(DiagonalBorderType.DiagonalDown), actualBorder.DiagonalDown);
        Assert.Equal(XLColor.Green, actualBorder.DiagonalBorderColor);
        Assert.Equal(XLBorderStyleValues.None, actualBorder.LeftBorder);
        Assert.Equal(XLBorderStyleValues.None, actualBorder.RightBorder);
        Assert.Equal(XLBorderStyleValues.None, actualBorder.TopBorder);
        Assert.Equal(XLBorderStyleValues.None, actualBorder.BottomBorder);
    }

    [Theory]
    [MemberData(nameof(TestData.StyledCellTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_MultipleBorders(CellType type)
    {
        // Arrange
        const string cellValue = "Border style test";

        var styles = new[] { BorderStyle.Thin, BorderStyle.DoubleLine, BorderStyle.DashDotDot, BorderStyle.Medium };
        var colors = new[] { Color.Firebrick, Color.ForestGreen, Color.Black, Color.Blue };

        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style
            {
                Border = new Border
                {
                    Left = new EdgeBorder { BorderStyle = styles[0], Color = colors[0] },
                    Right = new EdgeBorder { BorderStyle = styles[1], Color = colors[1] },
                    Top = new EdgeBorder { BorderStyle = styles[2], Color = colors[2] },
                    Bottom = new EdgeBorder { BorderStyle = styles[3], Color = colors[3] }
                }
            };
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        var actualBorder = actualCell.Style.Border;
        Assert.Equal(cellValue, actualCell.Value);

        Assert.Equal(styles[0].GetClosedXmlBorderStyle(), actualBorder.LeftBorder);
        Assert.Equal(styles[1].GetClosedXmlBorderStyle(), actualBorder.RightBorder);
        Assert.Equal(styles[2].GetClosedXmlBorderStyle(), actualBorder.TopBorder);
        Assert.Equal(styles[3].GetClosedXmlBorderStyle(), actualBorder.BottomBorder);

        Assert.Equal(XLColor.FromColor(colors[0]), actualBorder.LeftBorderColor);
        Assert.Equal(XLColor.FromColor(colors[1]), actualBorder.RightBorderColor);
        Assert.Equal(XLColor.FromColor(colors[2]), actualBorder.TopBorderColor);
        Assert.Equal(XLColor.FromColor(colors[3]), actualBorder.BottomBorderColor);

        Assert.Equal(XLBorderStyleValues.None, actualBorder.DiagonalBorder);
        Assert.Equal(XLColor.Black, actualBorder.DiagonalBorderColor);
    }

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_MultiFormatCellWithStringValue(bool formatting, CellType type)
    {
        // Arrange
        const string cellValue = "Formatting test";
        var fontName = formatting ? "Arial" : null;
        var fillColor = formatting ? Color.Green as Color? : null;
        var fontColor = formatting ? Color.Red as Color? : null;
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Fill.Color = fillColor;
            style.Font.Bold = formatting;
            style.Font.Color = fontColor;
            style.Font.Italic = formatting;
            style.Font.Name = fontName;
            style.Font.Strikethrough = formatting;
            style.NumberFormat = formatting ? NumberFormats.Percent : null;
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(cellValue, actualCell.Value);
        Assert.Equal(formatting, actualCell.Style.Fill.BackgroundColor.Color.ToArgb() == fillColor?.ToArgb());
        Assert.Equal(formatting, actualCell.Style.Font.Bold);
        Assert.Equal(formatting, actualCell.Style.Font.FontColor.Color.ToArgb() == fontColor?.ToArgb());
        Assert.Equal(formatting, string.Equals(actualCell.Style.Font.FontName, fontName, StringComparison.Ordinal));
        Assert.Equal(formatting, actualCell.Style.Font.Italic);
        Assert.Equal(formatting, actualCell.Style.Font.Strikethrough);
        Assert.Equal(formatting, actualCell.Style.NumberFormat.NumberFormatId == 9);
    }

    [Theory]
    [MemberData(nameof(TestData.StyledCellTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_MultipleCellsWithDifferentStyles(CellType type)
    {
        // Arrange
        var elements = new (string Value, Color FillColor, Color FontColor, string FontName, bool FontOption, double FontSize, string NumberFormat)[]
        {
            ("Value 1", Color.Blue, Color.PaleGoldenrod, "Times New Roman", true, 12, NumberFormats.Fraction),
            ("Value 2", Color.Snow, Color.Gainsboro, "Consolas", false, 20, "0.0000"),
            ("Value 3", Color.Aquamarine, Color.YellowGreen, "Impact", false, 18, "0.00000")
        };

        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var cells = elements.Select(x =>
            {
                var style = new Style
                {
                    Fill = { Color = x.FillColor },
                    Font =
                    {
                        Bold = x.FontOption,
                        Color = x.FontColor,
                        Italic = x.FontOption,
                        Name = x.FontName,
                        Size = x.FontSize,
                        Strikethrough = x.FontOption
                    },
                    NumberFormat = x.NumberFormat
                };

                var styleId = spreadsheet.AddStyle(style);
                return CellFactory.Create(type, x.Value, styleId);
            }).ToList();

            // Act
            await spreadsheet.AddRowAsync(cells);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var actualCells = workbook.Worksheets.Single().CellsUsed();
        Assert.All(actualCells, actualCell =>
        {
            var element = elements.Single(x => string.Equals(x.Value, actualCell.Value.ToString(), StringComparison.Ordinal));
            Assert.Equal(element.FillColor.ToArgb(), actualCell.Style.Fill.BackgroundColor.Color.ToArgb());
            Assert.Equal(element.FontColor.ToArgb(), actualCell.Style.Font.FontColor.Color.ToArgb());
            Assert.Equal(element.FontName, actualCell.Style.Font.FontName);
            Assert.Equal(element.FontOption, actualCell.Style.Font.Bold);
            Assert.Equal(element.FontOption, actualCell.Style.Font.Italic);
            Assert.Equal(element.FontOption, actualCell.Style.Font.Strikethrough);

            var numberFormatId = NumberFormats.GetPredefinedNumberFormatId(element.NumberFormat);
            if (numberFormatId is not null)
                Assert.Equal(numberFormatId, actualCell.Style.NumberFormat.NumberFormatId);
            else
                Assert.Equal(element.NumberFormat, actualCell.Style.NumberFormat.Format);
        });
    }

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_MixedCellTypeRows(bool firstRowStyled, CellType type)
    {
        // Arrange
        const string firstCellValue = "First";
        const string secondCellValue = "Second";
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Font.Bold = true;
            var styleId = spreadsheet.AddStyle(style);

            // Act
            if (firstRowStyled)
            {
                await spreadsheet.AddRowAsync(CellFactory.Create(type, firstCellValue, styleId));
                await spreadsheet.AddRowAsync(new DataCell(secondCellValue));
            }
            else
            {
                await spreadsheet.AddRowAsync(new DataCell(firstCellValue));
                await spreadsheet.AddRowAsync(CellFactory.Create(type, secondCellValue, styleId));
            }

            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualFirstCell = worksheet.Cell(1, 1);
        var actualSecondCell = worksheet.Cell(2, 1);
        Assert.Equal(firstCellValue, actualFirstCell.Value);
        Assert.Equal(secondCellValue, actualSecondCell.Value);
        Assert.Equal(firstRowStyled, actualFirstCell.Style.Font.Bold);
        Assert.Equal(!firstRowStyled, actualSecondCell.Style.Font.Bold);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Spreadsheet_AddRow_StyleChangeAfterAddStyleNotApplied(bool changedBefore)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Font.Bold = true;

            if (changedBefore) style.Font.Italic = true;
            var styleId = spreadsheet.AddStyle(style);
            if (!changedBefore) style.Font.Italic = true;

            var styledCell = new StyledCell("value", styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.True(actualCell.Style.Font.Bold);
        Assert.Equal(changedBefore, actualCell.Style.Font.Italic);
    }
}
