using ClosedXML.Excel;
using SpreadCheetah.Styling;
using SpreadCheetah.Test.Helpers;
using System.Drawing;
using Xunit;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetStyledRowTests
{
    public static IEnumerable<object?[]> TrueAndFalse => TestData.CombineWithStyledCellTypes(true, false);

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_BoldCellWithStringValue(bool bold, Type type)
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
    public async Task Spreadsheet_AddRow_SameBoldStyleCells(bool bold, Type type)
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
    public async Task Spreadsheet_AddRow_MixedBoldStyleCells(bool firstCellBold, Type type)
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
    public async Task Spreadsheet_AddRow_ItalicCellWithStringValue(bool italic, Type type)
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
    public async Task Spreadsheet_AddRow_StrikethroughCellWithStringValue(bool strikethrough, Type type)
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
    public async Task Spreadsheet_AddRow_FontSizeCellWithStringValue(double size, Type type)
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
    public async Task Spreadsheet_AddRow_FontColorCellWithStringValue(Type type)
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
        Assert.Equal(color, actualCell.Style.Font.FontColor.Color);
    }

    [Theory]
    [MemberData(nameof(TestData.StyledCellTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_FillColorCellWithStringValue(Type type)
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
        Assert.Equal(color, actualCell.Style.Fill.BackgroundColor.Color);
    }

    public static IEnumerable<object?[]> FontNames() => TestData.CombineWithStyledCellTypes(
        "Arial",
        "SimSun-ExtB",
        "Times New Roman",
        null);

    [Theory]
    [MemberData(nameof(FontNames))]
    public async Task Spreadsheet_AddRow_FontNameCellWithStringValue(string? fontName, Type type)
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
        NumberFormats.General,
        NumberFormats.NoDecimalPlaces,
        NumberFormats.PercentTwoDecimalPlaces,
        null);

    [Theory]
    [MemberData(nameof(PredefinedNumberFormats))]
    public async Task Spreadsheet_AddRow_PredefinedNumberFormatCellWithStringValue(string? numberFormat, Type type)
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
    public async Task Spreadsheet_AddRow_CustomNumberFormatCellWithStringValue(string numberFormat, Type type)
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
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_MultiFormatCellWithStringValue(bool formatting, Type type)
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
            style.NumberFormat = formatting ? NumberFormats.PercentInteger : null;
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
        Assert.Equal(formatting, actualCell.Style.Fill.BackgroundColor.Color == fillColor);
        Assert.Equal(formatting, actualCell.Style.Font.Bold);
        Assert.Equal(formatting, actualCell.Style.Font.FontColor.Color == fontColor);
        Assert.Equal(formatting, string.Equals(actualCell.Style.Font.FontName, fontName, StringComparison.Ordinal));
        Assert.Equal(formatting, actualCell.Style.Font.Italic);
        Assert.Equal(formatting, actualCell.Style.Font.Strikethrough);
        Assert.Equal(formatting, actualCell.Style.NumberFormat.NumberFormatId == 9);
    }

    [Theory]
    [MemberData(nameof(TrueAndFalse))]
    public async Task Spreadsheet_AddRow_MixedCellTypeRows(bool firstRowStyled, Type type)
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
}
