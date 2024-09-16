using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadCheetah.Styling;
using SpreadCheetah.Test.Helpers;
using System.Globalization;
using Alignment = SpreadCheetah.Styling.Alignment;
using Border = SpreadCheetah.Styling.Border;
using CellType = SpreadCheetah.Test.Helpers.CellType;
using Color = System.Drawing.Color;
using DiagonalBorder = SpreadCheetah.Styling.DiagonalBorder;
using Fill = SpreadCheetah.Styling.Fill;
using Font = SpreadCheetah.Styling.Font;
using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using SpreadsheetAssert = SpreadCheetah.TestHelpers.Assertions.SpreadsheetAssert;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetStyledRowTests
{
    [Theory]
    [MemberData(nameof(TestData.StyledCellAndValueTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_StyledCellWithValue(CellValueType valueType, bool isNull, CellType cellType, RowCollectionType rowType)
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
            await spreadsheet.AddRowAsync(styledCell, rowType);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(isNull, actualCell.Value.IsBlank);
        Assert.Equal(value?.ToString() ?? "", actualCell.Value.ToString(CultureInfo.CurrentCulture), ignoreCase: true);
        Assert.True(actualCell.Style.Font.Bold);
    }

    public static IEnumerable<object?[]> TrueAndFalse => TestData.CombineWithStyledCellTypes(true, false);

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_BoldCellWithStringValue(bool bold, CellType type, RowCollectionType rowType)
    {
        // Arrange
        const string cellValue = "Bold test";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        var style = new Style();
        style.Font.Bold = bold;
        var styleId = spreadsheet.AddStyle(style);
        var styledCell = CellFactory.Create(type, cellValue, styleId);

        // Act
        await spreadsheet.AddRowAsync(styledCell, rowType);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualCell = sheet["A1"];
        Assert.Equal(cellValue, actualCell.StringValue);
        Assert.Equal(bold, actualCell.Style.Font.Bold);
    }

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_SameBoldStyleCells(bool bold, CellType type, RowCollectionType rowType)
    {
        // Arrange
        const string firstCellValue = "First";
        const string secondCellValue = "Second";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        var style = new Style();
        style.Font.Bold = bold;
        var styleId = spreadsheet.AddStyle(style);

        var firstCell = CellFactory.Create(type, firstCellValue, styleId);
        var secondCell = CellFactory.Create(type, secondCellValue, styleId);

        // Act
        await spreadsheet.AddRowAsync([firstCell, secondCell], rowType);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualFirstCell = sheet["A1"];
        var actualSecondCell = sheet["B1"];
        Assert.Equal(firstCellValue, actualFirstCell.StringValue);
        Assert.Equal(secondCellValue, actualSecondCell.StringValue);
        Assert.Equal(bold, actualFirstCell.Style.Font.Bold);
        Assert.Equal(bold, actualSecondCell.Style.Font.Bold);
    }

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_MixedBoldStyleCells(bool firstCellBold, CellType type, RowCollectionType rowType)
    {
        // Arrange
        const string firstCellValue = "First";
        const string secondCellValue = "Second";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        var style = new Style();
        style.Font.Bold = true;
        var styleId = spreadsheet.AddStyle(style);

        var firstCell = CellFactory.Create(type, firstCellValue, firstCellBold ? styleId : null);
        var secondCell = CellFactory.Create(type, secondCellValue, firstCellBold ? null : styleId);

        // Act
        await spreadsheet.AddRowAsync([firstCell, secondCell], rowType);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualFirstCell = sheet["A1"];
        var actualSecondCell = sheet["B1"];
        Assert.Equal(firstCellValue, actualFirstCell.StringValue);
        Assert.Equal(secondCellValue, actualSecondCell.StringValue);
        Assert.Equal(firstCellBold, actualFirstCell.Style.Font.Bold);
        Assert.Equal(!firstCellBold, actualSecondCell.Style.Font.Bold);
    }

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_ItalicCellWithStringValue(bool italic, CellType type, RowCollectionType rowType)
    {
        // Arrange
        const string cellValue = "Italic test";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        var style = new Style();
        style.Font.Italic = italic;
        var styleId = spreadsheet.AddStyle(style);
        var styledCell = CellFactory.Create(type, cellValue, styleId);

        // Act
        await spreadsheet.AddRowAsync(styledCell, rowType);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualCell = sheet["A1"];
        Assert.Equal(cellValue, actualCell.StringValue);
        Assert.Equal(italic, actualCell.Style.Font.Italic);
    }

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_StrikethroughCellWithStringValue(bool strikethrough, CellType type, RowCollectionType rowType)
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
            await spreadsheet.AddRowAsync(styledCell, rowType);
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
    public async Task Spreadsheet_AddRow_FontSizeCellWithStringValue(double size, CellType type, RowCollectionType rowType)
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
            await spreadsheet.AddRowAsync(styledCell, rowType);
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
    public async Task Spreadsheet_AddRow_FontColorCellWithStringValue(CellType type, RowCollectionType rowType)
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
            await spreadsheet.AddRowAsync(styledCell, rowType);
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
    public async Task Spreadsheet_AddRow_FillColorCellWithStringValue(CellType type, RowCollectionType rowType)
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
            await spreadsheet.AddRowAsync(styledCell, rowType);
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
        "Name<With>Special&Chars",
        "NameThatIsExactly31CharsLong...",
        "<&><&><&><&><&><&><&><&><&><&>!",
        null);

    [Theory]
    [MemberData(nameof(FontNames))]
    public async Task Spreadsheet_AddRow_FontNameCellWithStringValue(string? fontName, CellType type, RowCollectionType rowType)
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
            await spreadsheet.AddRowAsync(styledCell, rowType);
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

    [Fact]
    public void Spreadsheet_AddRow_FontNameTooLong()
    {
        // Arrange
        const string fontName = "FontNameThatIsExactly32CharsLong";
        var style = new Style();

        // Act & Assert
        Assert.Throws<ArgumentException>(() => style.Font.Name = fontName);
    }

#pragma warning disable CS0618 // Type or member is obsolete - Testing for backwards compatibilty
    public static IEnumerable<object?[]> StandardNumberFormats() => TestData.CombineWithStyledCellTypes(
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
#pragma warning restore CS0618 // Type or member is obsolete

    [Theory]
    [MemberData(nameof(StandardNumberFormats))]
    public async Task Spreadsheet_AddRow_StandardNumberFormatCellWithStringValue(string? numberFormat, CellType type, RowCollectionType rowType)
    {
        // Arrange
        const string cellValue = "Number format test";
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

#pragma warning disable CS0618 // Type or member is obsolete - Testing legacy behaviour
            var style = new Style { NumberFormat = numberFormat };
#pragma warning restore CS0618 // Type or member is obsolete
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell, rowType);
            await spreadsheet.FinishAsync();
        }

        var expectedNumberFormatId = NumberFormatHelper.GetStandardNumberFormatId(numberFormat) ?? 0;

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
    public async Task Spreadsheet_AddRow_CustomNumberFormatCellWithStringValue(string numberFormat, CellType type, RowCollectionType rowType)
    {
        // Arrange
        const string cellValue = "Number format test";
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

#pragma warning disable CS0618 // Type or member is obsolete - Testing legacy behaviour
            var style = new Style { NumberFormat = numberFormat };
#pragma warning restore CS0618 // Type or member is obsolete
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell, rowType);
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
    public async Task Spreadsheet_AddRow_MultipleStylesWithTheSameCustomNumberFormat(CellType type, RowCollectionType rowType)
    {
        // Arrange
        const string cellValue = "Number format test";
        var format = NumberFormat.Custom("0.0000");
        var styles = new Style[]
        {
            new() { Format = format },
            new() { Format = format, Fill = new Fill { Color = Color.Coral } },
            new() { Format = format, Font = new Font { Bold = true } }
        };

        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var cells = styles.Select(x => CellFactory.Create(type, cellValue, spreadsheet.AddStyle(x))).ToList();

            // Act
            await spreadsheet.AddRowAsync(cells, rowType);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCells = worksheet.CellsUsed();
        Assert.All(actualCells, x => Assert.Equal(cellValue, x.Value));
        Assert.All(actualCells, x => Assert.Equal(format.ToString(), x.Style.NumberFormat.Format));
        Assert.Equal(styles.Length, actualCells.Count());
    }

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_DateTimeCell(bool withDefaultDateTimeFormat, CellType type, RowCollectionType rowType)
    {
        // Arrange
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        if (!withDefaultDateTimeFormat)
            options.DefaultDateTimeFormat = null;

        var value = new DateTime(2021, 2, 3, 4, 5, 6, DateTimeKind.Unspecified);
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        await spreadsheet.StartWorksheetAsync("Sheet");

        var style = new Style { Font = { Italic = true } };
        var styleId = spreadsheet.AddStyle(style);
        var styledCell = CellFactory.Create(type, value, styleId);
        var expectedNumberFormat = withDefaultDateTimeFormat ? @"yyyy\-mm\-dd\ hh:mm:ss" : "";

        // Act
        await spreadsheet.AddRowAsync(styledCell, rowType);
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(value.ToOADate(), actualCell.Value.GetUnifiedNumber(), 0.0005);
        Assert.Equal(expectedNumberFormat, actualCell.Style.NumberFormat.Format);
        Assert.True(actualCell.Style.Font.Italic);
    }

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_DateTimeNumberFormat(bool withExplicitNumberFormat, CellType type, RowCollectionType rowType)
    {
        // Arrange
        const string explicitNumberFormat = "yyyy";
        var expectedNumberFormat = withExplicitNumberFormat ? explicitNumberFormat : NumberFormats.DateTimeSortable;
        var value = new DateTime(2021, 2, 3, 4, 5, 6, DateTimeKind.Unspecified);
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Font.Italic = true;
            if (withExplicitNumberFormat)
                style.Format = NumberFormat.Custom(explicitNumberFormat);

            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, value, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell, rowType);
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

    public static IEnumerable<object?[]> BorderStyles() => TestData.CombineWithStyledCellTypes(EnumPolyfill.GetValues<BorderStyle>());

    [Theory]
    [MemberData(nameof(BorderStyles))]
    public async Task Spreadsheet_AddRow_BorderStyle(BorderStyle borderStyle, CellType type, RowCollectionType rowType)
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
            await spreadsheet.AddRowAsync(styledCell, rowType);
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

    public static IEnumerable<object?[]> DiagonalBorderTypes() => TestData.CombineWithStyledCellTypes(EnumPolyfill.GetValues<DiagonalBorderType>());

    [Theory]
    [MemberData(nameof(DiagonalBorderTypes))]
    public async Task Spreadsheet_AddRow_DiagonalBorder(DiagonalBorderType borderType, CellType type, RowCollectionType rowType)
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
            await spreadsheet.AddRowAsync(styledCell, rowType);
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
    public async Task Spreadsheet_AddRow_MultipleBorders(CellType type, RowCollectionType rowType)
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
            await spreadsheet.AddRowAsync(styledCell, rowType);
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

    public static IEnumerable<object?[]> HorizontalAlignments() => TestData.CombineWithStyledCellTypes(EnumPolyfill.GetValues<HorizontalAlignment>());

    [Theory]
    [MemberData(nameof(HorizontalAlignments))]
    public async Task Spreadsheet_AddRow_HorizontalAlignment(HorizontalAlignment alignment, CellType type, RowCollectionType rowType)
    {
        // Arrange
        const string cellValue = "Alignment test";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var style = new Style();
        style.Alignment.Horizontal = alignment;
        var styleId = spreadsheet.AddStyle(style);
        var styledCell = CellFactory.Create(type, cellValue, styleId);
        var expectedAlignment = alignment.GetClosedXmlHorizontalAlignment();

        // Act
        await spreadsheet.AddRowAsync(styledCell, rowType);
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        var actualAlignment = actualCell.Style.Alignment;
        Assert.Equal(cellValue, actualCell.Value);
        Assert.Equal(expectedAlignment, actualAlignment.Horizontal);
        Assert.Equal(XLAlignmentVerticalValues.Bottom, actualAlignment.Vertical);
        Assert.Equal(0, actualAlignment.Indent);
        Assert.False(actualAlignment.WrapText);
    }

    public static IEnumerable<object?[]> VerticalAlignments() => TestData.CombineWithStyledCellTypes(EnumPolyfill.GetValues<VerticalAlignment>());

    [Theory]
    [MemberData(nameof(VerticalAlignments))]
    public async Task Spreadsheet_AddRow_VerticalAlignment(VerticalAlignment alignment, CellType type, RowCollectionType rowType)
    {
        // Arrange
        const string cellValue = "Alignment test";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var style = new Style();
        style.Alignment.Vertical = alignment;
        var styleId = spreadsheet.AddStyle(style);
        var styledCell = CellFactory.Create(type, cellValue, styleId);
        var expectedAlignment = alignment.GetClosedXmlVerticalAlignment();

        // Act
        await spreadsheet.AddRowAsync(styledCell, rowType);
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        var actualAlignment = actualCell.Style.Alignment;
        Assert.Equal(cellValue, actualCell.Value);
        Assert.Equal(XLAlignmentHorizontalValues.General, actualAlignment.Horizontal);
        Assert.Equal(expectedAlignment, actualAlignment.Vertical);
        Assert.Equal(0, actualAlignment.Indent);
        Assert.False(actualAlignment.WrapText);
    }

    [Theory]
    [MemberData(nameof(TestData.StyledCellTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_MixedAlignment(CellType type, RowCollectionType rowType)
    {
        // Arrange
        const string cellValue = "Alignment test";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var style = new Style
        {
            Alignment = new Alignment
            {
                Horizontal = HorizontalAlignment.Right,
                Indent = 2,
                Vertical = VerticalAlignment.Center,
                WrapText = true
            }
        };

        var styleId = spreadsheet.AddStyle(style);
        var styledCell = CellFactory.Create(type, cellValue, styleId);

        // Act
        await spreadsheet.AddRowAsync(styledCell, rowType);
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        var actualAlignment = actualCell.Style.Alignment;
        Assert.Equal(cellValue, actualCell.Value);
        Assert.Equal(XLAlignmentHorizontalValues.Right, actualAlignment.Horizontal);
        Assert.Equal(XLAlignmentVerticalValues.Center, actualAlignment.Vertical);
        Assert.Equal(2, actualAlignment.Indent);
        Assert.True(actualAlignment.WrapText);
    }

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_MultiFormatCellWithStringValue(bool formatting, CellType type, RowCollectionType rowType)
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
            style.Format = formatting ? NumberFormat.Standard(StandardNumberFormat.Percent) : null;
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell, rowType);
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
    public async Task Spreadsheet_AddRow_MultipleCellsWithDifferentStyles(CellType type, RowCollectionType rowType)
    {
        // Arrange
        var elements = new (string Value, Color FillColor, Color FontColor, string FontName, bool FontOption, double FontSize, NumberFormat Format)[]
        {
            ("Value 1", Color.Blue, Color.PaleGoldenrod, "Times New Roman", true, 12, NumberFormat.Standard(StandardNumberFormat.Fraction)),
            ("Value 2", Color.Snow, Color.Gainsboro, "Consolas", false, 20, NumberFormat.Custom("0.0000")),
            ("Value 3", Color.Aquamarine, Color.YellowGreen, "Impact", false, 18, NumberFormat.Custom("0.0000"))
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
                    Format = x.Format
                };

                var styleId = spreadsheet.AddStyle(style);
                return CellFactory.Create(type, x.Value, styleId);
            }).ToList();

            // Act
            await spreadsheet.AddRowAsync(cells, rowType);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var actualCells = workbook.Worksheets.Single().CellsUsed();
        Assert.All(actualCells, actualCell =>
        {
            var element = elements.Single(x => string.Equals(x.Value, actualCell.Value.GetText(), StringComparison.Ordinal));
            Assert.Equal(element.FillColor.ToArgb(), actualCell.Style.Fill.BackgroundColor.Color.ToArgb());
            Assert.Equal(element.FontColor.ToArgb(), actualCell.Style.Font.FontColor.Color.ToArgb());
            Assert.Equal(element.FontName, actualCell.Style.Font.FontName);
            Assert.Equal(element.FontOption, actualCell.Style.Font.Bold);
            Assert.Equal(element.FontOption, actualCell.Style.Font.Italic);
            Assert.Equal(element.FontOption, actualCell.Style.Font.Strikethrough);

            if (element.Format.Equals(NumberFormat.Standard(StandardNumberFormat.Fraction)))
                Assert.Equal((int)StandardNumberFormat.Fraction, actualCell.Style.NumberFormat.NumberFormatId);
            else
                Assert.Equal(element.Format.ToString(), actualCell.Style.NumberFormat.Format);
        });
    }

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_MixedCellTypeRows(bool firstRowStyled, CellType type, RowCollectionType rowType)
    {
        // Arrange
        const string firstCellValue = "First";
        const string secondCellValue = "Second";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        var style = new Style();
        style.Font.Bold = true;
        var styleId = spreadsheet.AddStyle(style);

        // Act
        if (firstRowStyled)
        {
            await spreadsheet.AddRowAsync(CellFactory.Create(type, firstCellValue, styleId), rowType);
            await spreadsheet.AddRowAsync(new DataCell(secondCellValue), rowType);
        }
        else
        {
            await spreadsheet.AddRowAsync(new DataCell(firstCellValue), rowType);
            await spreadsheet.AddRowAsync(CellFactory.Create(type, secondCellValue, styleId), rowType);
        }

        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualFirstCell = sheet["A1"];
        var actualSecondCell = sheet["A2"];
        Assert.Equal(firstCellValue, actualFirstCell.StringValue);
        Assert.Equal(secondCellValue, actualSecondCell.StringValue);
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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
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

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualCell = sheet["A1"];
        Assert.True(actualCell.Style.Font.Bold);
        Assert.Equal(changedBefore, actualCell.Style.Font.Italic);
    }

    [Fact]
    public async Task Spreadsheet_AddRow_MultipleStyles()
    {
        // Arrange
        const int count = 1000;
        const string cellValue = "Fill test";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize });
        await spreadsheet.StartWorksheetAsync("Sheet");
        var styles = StyleGenerator.Generate(count);
        var uniqueStyleIds = new HashSet<int>();

        foreach (var style in styles)
        {
            var styleId = spreadsheet.AddStyle(style);
            uniqueStyleIds.Add(styleId.Id);

            // Act
            await spreadsheet.AddRowAsync(new StyledCell(cellValue, styleId));
        }

        await spreadsheet.FinishAsync();

        // Assert
        Assert.True(uniqueStyleIds.Count > count * .8f);
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualStyles = worksheet.Cells().Select(x => x.Style).ToList();
        Assert.All(styles.Zip(actualStyles), x => SpreadCheetah.Test.Helpers.SpreadsheetAssert.EquivalentStyle(x.First, x.Second));
    }

    [Theory]
    [MemberData(nameof(TestData.StyledCellAndValueTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_ExplicitCellReferencesForStyledCells(CellValueType valueType, bool isNull, CellType cellType, RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize, WriteCellReferenceAttributes = true };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        var style = new Style();
        style.Fill.Color = Color.LightSeaGreen;
        var styleId = spreadsheet.AddStyle(style);

        var row1 = Enumerable.Range(1, 10).Select(_ => CellFactory.Create(cellType, valueType, isNull, styleId, out var value)).ToList();
        var row2 = Enumerable.Range(1, 1).Select(_ => CellFactory.Create(cellType, valueType, isNull, styleId, out var value)).ToList();
        var row3 = Enumerable.Range(1, 100).Select(_ => CellFactory.Create(cellType, valueType, isNull, styleId, out var value)).ToList();

        var expectedRow1Refs = Enumerable.Range(1, 10).Select(x => SpreadsheetUtility.GetColumnName(x) + "1").OfType<string?>();
        var expectedRow2Refs = Enumerable.Range(1, 1).Select(x => SpreadsheetUtility.GetColumnName(x) + "2").OfType<string?>();
        var expectedRow3Refs = Enumerable.Range(1, 100).Select(x => SpreadsheetUtility.GetColumnName(x) + "3").OfType<string?>();

        // Act
        await spreadsheet.StartWorksheetAsync("Sheet1");
        await spreadsheet.AddRowAsync(row1, rowType);
        await spreadsheet.AddRowAsync(row2, rowType);
        await spreadsheet.AddRowAsync(row3, rowType);

        await spreadsheet.StartWorksheetAsync("Sheet2");
        await spreadsheet.AddRowAsync(row1, rowType);

        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetParts = actual.WorkbookPart!.WorksheetParts.ToList();
        Assert.Equal(2, sheetParts.Count);

        var sheet1Rows = sheetParts[0].Worksheet.Descendants<Row>().ToList();
        Assert.Equal(3, sheet1Rows.Count);

        var actualRow1Refs = sheet1Rows[0].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        var actualRow2Refs = sheet1Rows[1].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        var actualRow3Refs = sheet1Rows[2].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);

        Assert.Equal(expectedRow1Refs, actualRow1Refs);
        Assert.Equal(expectedRow2Refs, actualRow2Refs);
        Assert.Equal(expectedRow3Refs, actualRow3Refs);

        var sheet2Row = sheetParts[1].Worksheet.Descendants<Row>().Single();
        var actualSheet2Refs = sheet2Row.Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        Assert.Equal(expectedRow1Refs, actualSheet2Refs);
    }

    [Theory]
    [MemberData(nameof(TestData.StyledCellTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_LongStringValueStyledCells(CellType cellType, RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        var value = new string('a', options.BufferSize * 2);
        var style = new Style();
        style.Fill.Color = Color.Thistle;
        var styleId = spreadsheet.AddStyle(style);

        var row1 = Enumerable.Range(1, 10).Select(_ => CellFactory.Create(cellType, value, styleId)).ToList();
        var row2 = Enumerable.Range(1, 1).Select(_ => CellFactory.Create(cellType, value, styleId)).ToList();
        var row3 = Enumerable.Range(1, 100).Select(_ => CellFactory.Create(cellType, value, styleId)).ToList();

        // Act
        await spreadsheet.StartWorksheetAsync("Sheet1");
        await spreadsheet.AddRowAsync(row1, rowType);
        await spreadsheet.AddRowAsync(row2, rowType);
        await spreadsheet.AddRowAsync(row3, rowType);

        await spreadsheet.StartWorksheetAsync("Sheet2");
        await spreadsheet.AddRowAsync(row1, rowType);

        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetParts = actual.WorkbookPart!.WorksheetParts.ToList();
        Assert.Equal(2, sheetParts.Count);

        var sheet1Rows = sheetParts[0].Worksheet.Descendants<Row>().ToList();
        Assert.Equal(3, sheet1Rows.Count);
        Assert.All(sheet1Rows[0].Descendants<OpenXmlCell>(), x => Assert.Equal(value, x.InnerText));
        Assert.All(sheet1Rows[1].Descendants<OpenXmlCell>(), x => Assert.Equal(value, x.InnerText));
        Assert.All(sheet1Rows[2].Descendants<OpenXmlCell>(), x => Assert.Equal(value, x.InnerText));

        var sheet2Row = Assert.Single(sheetParts[1].Worksheet.Descendants<Row>());
        Assert.All(sheet2Row.Descendants<OpenXmlCell>(), x => Assert.Equal(value, x.InnerText));
    }

    [Theory]
    [MemberData(nameof(TestData.StyledCellTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_ExplicitCellReferencesForLongStringValueStyledCells(CellType cellType, RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize, WriteCellReferenceAttributes = true };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        var value = new string('a', options.BufferSize * 2);
        var style = new Style();
        style.Fill.Color = Color.Thistle;
        var styleId = spreadsheet.AddStyle(style);

        var row1 = Enumerable.Range(1, 10).Select(_ => CellFactory.Create(cellType, value, styleId)).ToList();
        var row2 = Enumerable.Range(1, 1).Select(_ => CellFactory.Create(cellType, value, styleId)).ToList();
        var row3 = Enumerable.Range(1, 100).Select(_ => CellFactory.Create(cellType, value, styleId)).ToList();

        var expectedRow1Refs = Enumerable.Range(1, 10).Select(x => SpreadsheetUtility.GetColumnName(x) + "1").OfType<string?>();
        var expectedRow2Refs = Enumerable.Range(1, 1).Select(x => SpreadsheetUtility.GetColumnName(x) + "2").OfType<string?>();
        var expectedRow3Refs = Enumerable.Range(1, 100).Select(x => SpreadsheetUtility.GetColumnName(x) + "3").OfType<string?>();

        // Act
        await spreadsheet.StartWorksheetAsync("Sheet1");
        await spreadsheet.AddRowAsync(row1, rowType);
        await spreadsheet.AddRowAsync(row2, rowType);
        await spreadsheet.AddRowAsync(row3, rowType);

        await spreadsheet.StartWorksheetAsync("Sheet2");
        await spreadsheet.AddRowAsync(row1, rowType);

        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetParts = actual.WorkbookPart!.WorksheetParts.ToList();
        Assert.Equal(2, sheetParts.Count);

        var sheet1Rows = sheetParts[0].Worksheet.Descendants<Row>().ToList();
        Assert.Equal(3, sheet1Rows.Count);

        var actualRow1Refs = sheet1Rows[0].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        var actualRow2Refs = sheet1Rows[1].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        var actualRow3Refs = sheet1Rows[2].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);

        Assert.Equal(expectedRow1Refs, actualRow1Refs);
        Assert.Equal(expectedRow2Refs, actualRow2Refs);
        Assert.Equal(expectedRow3Refs, actualRow3Refs);

        var sheet2Row = sheetParts[1].Worksheet.Descendants<Row>().Single();
        var actualSheet2Refs = sheet2Row.Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        Assert.Equal(expectedRow1Refs, actualSheet2Refs);
    }

    public static IEnumerable<object?[]> ExplicitStandardNumberFormats() => TestData.CombineWithStyledCellTypes(EnumPolyfill.GetValues<StandardNumberFormat>());

    [Theory]
    [MemberData(nameof(ExplicitStandardNumberFormats))]
    public async Task Spreadsheet_AddRow_StandardNumberFormatCellExplicitly(StandardNumberFormat numberFormat, CellType type, RowCollectionType rowType)
    {
        // Arrange
        const string cellValue = "Number format test";
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style { Format = NumberFormat.Standard(numberFormat) };
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell, rowType);
            await spreadsheet.FinishAsync();
        }

        var expectedNumberFormatId = (int)numberFormat;

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(cellValue, actualCell.Value);
        Assert.Equal(expectedNumberFormatId, actualCell.Style.NumberFormat.NumberFormatId);
    }

#pragma warning disable CS0618 // Type or member is obsolete - Testing legacy behaviour
    public static IEnumerable<object?[]> CustomNumberFormatsMatchingStandardFormats() => TestData.CombineWithStyledCellTypes(
        NumberFormats.General,
        NumberFormats.Fraction,
        NumberFormats.FractionTwoDenominatorPlaces,
        NumberFormats.NoDecimalPlaces,
        NumberFormats.Percent,
        NumberFormats.PercentTwoDecimalPlaces,
        NumberFormats.Scientific,
        NumberFormats.ThousandsSeparator,
        NumberFormats.ThousandsSeparatorTwoDecimalPlaces,
        NumberFormats.TwoDecimalPlaces,
        "mm-dd-yy",
        "d-mmm-yy",
        "d-mmm",
        "mmm-yy",
        "h:mm AM/PM",
        "h:mm:ss AM/PM",
        "h:mm",
        "h:mm:ss",
        "m/d/yy h:mm",
        "#,##0 ;(#,##0)",
        "#,##0 ;[Red](#,##0)",
        "#,##0.00;(#,##0.00)",
        "#,##0.00;[Red](#,##0.00)",
        "mm:ss",
        "[h]:mm:ss",
        "mmss.0",
        "##0.0E+0",
        NumberFormats.Text);
#pragma warning restore CS0618 // Type or member is obsolete

    [Theory]
    [MemberData(nameof(CustomNumberFormats))]
    [MemberData(nameof(CustomNumberFormatsMatchingStandardFormats))] // Check that old hardcoded standard number formats can be explictly specified as custom formats
    public async Task Spreadsheet_AddRow_CustomNumberFormatCellWithStringValueExplicitly(string numberFormat, CellType type, RowCollectionType rowType)
    {
        // Arrange
        const string cellValue = "Number format test";
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style { Format = NumberFormat.Custom(numberFormat) };
            var styleId = spreadsheet.AddStyle(style);
            var styledCell = CellFactory.Create(type, cellValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(styledCell, rowType);
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
}
