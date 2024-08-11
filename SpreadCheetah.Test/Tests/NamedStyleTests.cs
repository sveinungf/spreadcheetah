using OfficeOpenXml;
using SpreadCheetah.Styling;
using SpreadCheetah.Test.Helpers;
using System.Drawing;

namespace SpreadCheetah.Test.Tests;

public class NamedStyleTests
{
    private static readonly SpreadCheetahOptions SpreadCheetahOptions = new() { BufferSize = SpreadCheetahOptions.MinimumBufferSize };

    [Theory]
    [InlineData(StyleNameVisibility.Visible)]
    [InlineData(StyleNameVisibility.Hidden)]
    public async Task Spreadsheet_AddStyle_NamedStyle(StyleNameVisibility visibility)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, SpreadCheetahOptions);
        await spreadsheet.StartWorksheetAsync("My sheet");
        var style = new Style { Font = { Bold = true } };
        const string name = "My bold style";

        // Act
        var addedStyleId = spreadsheet.AddStyle(style, name, visibility);
        var returnedStyleId = spreadsheet.GetStyleId(name);
        await spreadsheet.FinishAsync();

        // Assert
        Assert.Equal(addedStyleId, returnedStyleId);
        SpreadsheetAssert.Valid(stream);
        using var package = new ExcelPackage(stream);
        var namedStyles = package.Workbook.Styles.NamedStyles;
        var namedStyle = Assert.Single(namedStyles, x => !x.Name.Equals("Normal", StringComparison.Ordinal));
        Assert.Equal(name, namedStyle.Name);
        Assert.True(namedStyle.Style.Font.Bold);
    }

    [Fact]
    public async Task Spreadsheet_AddStyle_NamedStyleWithNoVisiblity()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, SpreadCheetahOptions);
        await spreadsheet.StartWorksheetAsync("My sheet");
        var style = new Style { Font = { Bold = true } };
        const string name = "My bold style";

        // Act
        var addedStyleId = spreadsheet.AddStyle(style, name);
        var returnedStyleId = spreadsheet.GetStyleId(name);
        await spreadsheet.FinishAsync();

        // Assert
        Assert.Equal(addedStyleId, returnedStyleId);
        SpreadsheetAssert.Valid(stream);
        using var package = new ExcelPackage(stream);
        var namedStyle = Assert.Single(package.Workbook.Styles.NamedStyles);
        Assert.Equal("Normal", namedStyle.Name);
        Assert.False(namedStyle.Style.Font.Bold);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_AddStyle_NamedStyleUsedByCell(bool withDefaultDateTimeFormat)
    {
        // Arrange
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        if (!withDefaultDateTimeFormat)
            options.DefaultDateTimeFormat = null;

        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        await spreadsheet.StartWorksheetAsync("My sheet");
        var style = new Style { Font = { Bold = true } };
        const string name = "My bold style";

        // Act
        var styleId = spreadsheet.AddStyle(style, name, StyleNameVisibility.Visible);
        await spreadsheet.AddRowAsync([new StyledCell("My cell", styleId)]);
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var package = new ExcelPackage(stream);
        var worksheet = Assert.Single(package.Workbook.Worksheets);
        var actualCell = Assert.Single(worksheet.Cells);
        Assert.Equal(name, actualCell.StyleName);
        Assert.True(actualCell.Style.Font.Bold);
    }

    [Fact]
    public async Task Spreadsheet_AddStyle_NamedStyleUsedByMultipleCells()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, SpreadCheetahOptions);
        var style = new Style { Font = { Bold = true } };
        const string name = "My bold style";

        // Act
        var styleId = spreadsheet.AddStyle(style, name, StyleNameVisibility.Visible);
        await spreadsheet.StartWorksheetAsync("Sheet 1");
        await spreadsheet.AddRowAsync([new StyledCell("A1", styleId), new StyledCell(2, null), new StyledCell(3, styleId)]);
        await spreadsheet.AddRowAsync([new StyledCell("A2", styleId)]);
        await spreadsheet.StartWorksheetAsync("Sheet 2");
        await spreadsheet.AddRowAsync([new StyledCell("A1", styleId)]);
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var package = new ExcelPackage(stream);
        var worksheet1 = package.Workbook.Worksheets["Sheet 1"];
        Assert.Equal(name, worksheet1.Cells["A1"].StyleName);
        Assert.NotEqual(name, worksheet1.Cells["B2"].StyleName);
        Assert.Equal(name, worksheet1.Cells["C1"].StyleName);
        Assert.Equal(name, worksheet1.Cells["A2"].StyleName);

        var worksheet2 = package.Workbook.Worksheets["Sheet 2"];
        Assert.Equal(name, worksheet2.Cells["A1"].StyleName);
    }

    [Fact]
    public async Task Spreadsheet_AddStyle_MultipleNamedStylesUsedByMultipleCells()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, SpreadCheetahOptions);
        (string Name, Style Style)[] styles =
        [
            ("Bold", new Style { Font = { Bold = true } }),
            ("Italic", new Style { Font = { Italic = true } }),
            ("Red", new Style { Fill = { Color = Color.Red } }),
            ("Blue", new Style { Fill = { Color = Color.Blue } })
        ];

        // Act
        var styleIds = styles.Select(x => spreadsheet.AddStyle(x.Style, x.Name, StyleNameVisibility.Visible)).ToList();

        await spreadsheet.StartWorksheetAsync("Sheet 1");
        await spreadsheet.AddRowAsync([new StyledCell("A1", styleIds[0]), new StyledCell(2, null), new StyledCell(3, styleIds[1])]);
        await spreadsheet.AddRowAsync([new StyledCell("A2", styleIds[2])]);
        await spreadsheet.StartWorksheetAsync("Sheet 2");
        await spreadsheet.AddRowAsync([new StyledCell("A1", styleIds[3])]);
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var package = new ExcelPackage(stream);
        var worksheet1 = package.Workbook.Worksheets["Sheet 1"];
        Assert.Equal(styles[0].Name, worksheet1.Cells["A1"].StyleName);
        Assert.Equal(styles[1].Name, worksheet1.Cells["C1"].StyleName);
        Assert.Equal(styles[2].Name, worksheet1.Cells["A2"].StyleName);

        var worksheet2 = package.Workbook.Worksheets["Sheet 2"];
        Assert.Equal(styles[3].Name, worksheet2.Cells["A1"].StyleName);
    }

    [Fact]
    public async Task Spreadsheet_AddStyle_NamedStyleWithInvalidVisibility()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, SpreadCheetahOptions);
        var style = new Style { Font = { Bold = true } };

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => spreadsheet.AddStyle(style, "Name", (StyleNameVisibility)3));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData(" Style")]
    [InlineData("Style ")]
    [InlineData("Normal")]
    public async Task Spreadsheet_AddStyle_NamedStyleWithInvalidName(string? name)
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, SpreadCheetahOptions);
        var style = new Style { Font = { Bold = true } };

        // Act & Assert
        Assert.ThrowsAny<ArgumentException>(() => spreadsheet.AddStyle(style, name!));
    }

    [Fact]
    public async Task Spreadsheet_AddStyle_NamedStyleWithTooLongName()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, SpreadCheetahOptions);
        var style = new Style { Font = { Bold = true } };
        var name = new string('c', 256);

        // Act & Assert
        Assert.Throws<ArgumentException>(() => spreadsheet.AddStyle(style, name));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Spreadsheet_AddStyle_NamedStyleWithDuplicateName(bool differentCasing)
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, SpreadCheetahOptions);

        var style = new Style { Font = { Bold = true } };
        const string name = "My bold style";
        spreadsheet.AddStyle(style, name);

        var otherStyle = new Style { Font = { Italic = true } };
        var duplicateName = differentCasing ? name.ToUpperInvariant() : name;

        // Act & Assert
        Assert.Throws<ArgumentException>(() => spreadsheet.AddStyle(otherStyle, duplicateName));
    }

    [Fact]
    public async Task Spreadsheet_AddStyle_DuplicateNamedStylesReturnsDifferentStyleIds()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, SpreadCheetahOptions);
        var style1 = new Style { Font = { Bold = true } };
        var style2 = style1 with { };

        // Act
        var styleId1 = spreadsheet.AddStyle(style1, "Style 1");
        var styleId2 = spreadsheet.AddStyle(style2, "Style 2");

        // Assert
        Assert.NotEqual(styleId1.Id, styleId2.Id);
    }

    [Fact]
    public async Task Spreadsheet_AddStyle_NamedStyleDuplicateOfUnnamedStyleReturnsDifferentStyleIds()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, SpreadCheetahOptions);
        var style1 = new Style { Font = { Bold = true } };
        var style2 = style1 with { };

        // Act
        var styleId1 = spreadsheet.AddStyle(style1);
        var styleId2 = spreadsheet.AddStyle(style2, "Style 2");

        // Assert
        Assert.NotEqual(styleId1.Id, styleId2.Id);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Spreadsheet_GetStyleId_IncorrectName(bool existingNamedStyle)
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, SpreadCheetahOptions);

        if (existingNamedStyle)
        {
            var style = new Style { Font = { Bold = true } };
            const string name = "My bold style";
            spreadsheet.AddStyle(style, name);
        }

        // Act & Assert
        Assert.Throws<SpreadCheetahException>(() => spreadsheet.GetStyleId("Other style"));
    }
}
