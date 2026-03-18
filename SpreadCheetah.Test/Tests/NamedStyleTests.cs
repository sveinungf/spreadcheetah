using SpreadCheetah.Styling;
using SpreadCheetah.Test.Extensions;
using SpreadCheetah.Test.Helpers;
using System.Drawing;
using System.IO.Compression;
using SpreadsheetAssert = SpreadCheetah.TestHelpers.Assertions.SpreadsheetAssert;

namespace SpreadCheetah.Test.Tests;

public class NamedStyleTests
{
    private static readonly SpreadCheetahOptions SpreadCheetahOptions = new() { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Theory]
    [InlineData(StyleNameVisibility.Visible)]
    [InlineData(StyleNameVisibility.Hidden)]
    public async Task Spreadsheet_AddStyle_NamedStyle(StyleNameVisibility visibility)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, SpreadCheetahOptions, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new Style { Font = { Bold = true } };
        const string name = "My bold style";

        // Act
        var addedStyleId = spreadsheet.AddStyle(style, name, visibility);
        var returnedStyleId = spreadsheet.GetStyleId(name);
        await spreadsheet.FinishAsync(Token);

        // Assert
        Assert.Equal(addedStyleId, returnedStyleId);
        SpreadsheetAssert.Valid(stream);
        using var zip = await ZipArchive.CreateAsync(stream, Token);
        using var stylesXml = await zip.GetStylesXmlStreamAsync(Token);
        await VerifyXml(stylesXml);
    }

    [Fact]
    public async Task Spreadsheet_AddStyle_NamedStyleWithCharacterToEscape()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, SpreadCheetahOptions, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new Style { Font = { Bold = true } };
        const string name = "Nice & cool style";

        // Act
        spreadsheet.AddStyle(style, name, StyleNameVisibility.Visible);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var zip = await ZipArchive.CreateAsync(stream, Token);
        using var stylesXml = await zip.GetStylesXmlStreamAsync(Token);
        await VerifyXml(stylesXml);
    }

    [Fact]
    public async Task Spreadsheet_AddStyle_NamedStyleWithNoVisiblity()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, SpreadCheetahOptions, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new Style { Font = { Bold = true } };
        const string name = "My bold style";

        // Act
        var addedStyleId = spreadsheet.AddStyle(style, name);
        var returnedStyleId = spreadsheet.GetStyleId(name);
        await spreadsheet.FinishAsync(Token);

        // Assert
        Assert.Equal(addedStyleId, returnedStyleId);
        SpreadsheetAssert.Valid(stream);
        using var zip = await ZipArchive.CreateAsync(stream, Token);
        using var stylesXml = await zip.GetStylesXmlStreamAsync(Token);
        await VerifyXml(stylesXml);
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

        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new Style { Font = { Bold = true } };
        const string name = "My bold style";

        // Act
        var styleId = spreadsheet.AddStyle(style, name, StyleNameVisibility.Visible);
        await spreadsheet.AddRowAsync([new StyledCell("My cell", styleId)], token: Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.True(sheet["A1"].Style.Font.Bold);
    }

    [Fact]
    public async Task Spreadsheet_AddStyle_NamedStyleUsedByMultipleCells()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, SpreadCheetahOptions, Token);
        var style = new Style { Font = { Bold = true } };
        const string name = "My bold style";

        // Act
        var styleId = spreadsheet.AddStyle(style, name, StyleNameVisibility.Visible);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        await spreadsheet.AddRowAsync([new StyledCell("A1", styleId), new StyledCell(2, null), new StyledCell(3, styleId)], token: Token);
        await spreadsheet.AddRowAsync([new StyledCell("A2", styleId)], token: Token);
        await spreadsheet.StartWorksheetAsync("Sheet 2", token: Token);
        await spreadsheet.AddRowAsync([new StyledCell("A1", styleId)], token: Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheets = SpreadsheetAssert.Sheets(stream);
        Assert.True(sheets[0]["A1"].Style.Font.Bold);
        Assert.True(sheets[1]["A1"].Style.Font.Bold);
        using var zip = await ZipArchive.CreateAsync(stream, Token);
        using var stylesXml = await zip.GetStylesXmlStreamAsync(Token);
        await VerifyXml(stylesXml);
    }

    [Fact]
    public async Task Spreadsheet_AddStyle_MultipleNamedStylesUsedByMultipleCells()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, SpreadCheetahOptions, Token);
        (string Name, Style Style)[] styles =
        [
            ("Bold", new Style { Font = { Bold = true } }),
            ("Red", new Style { Fill = { Color = Color.Red } }),
            ("Blue", new Style { Fill = { Color = Color.Blue } }),
            ("Italic", new Style { Font = { Italic = true } })
        ];

        // Act
        var styleIds = styles.Select(x => spreadsheet.AddStyle(x.Style, x.Name, StyleNameVisibility.Visible)).ToList();

        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        await spreadsheet.AddRowAsync([new StyledCell("A1", styleIds[0]), new StyledCell(2, null), new StyledCell(3, styleIds[1])], token: Token);
        await spreadsheet.AddRowAsync([new StyledCell("A2", styleIds[2])], token: Token);
        await spreadsheet.StartWorksheetAsync("Sheet 2", token: Token);
        await spreadsheet.AddRowAsync([new StyledCell("A1", styleIds[3])], token: Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheets = SpreadsheetAssert.Sheets(stream);
        Assert.True(sheets[0]["A1"].Style.Font.Bold);
        Assert.True(sheets[1]["A1"].Style.Font.Italic);
        using var zip = await ZipArchive.CreateAsync(stream, Token);
        using var stylesXml = await zip.GetStylesXmlStreamAsync(Token);
        await VerifyXml(stylesXml);
    }

    [Fact]
    public async Task Spreadsheet_AddStyle_MultipleNamedStylesWithLongNames()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, SpreadCheetahOptions, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var style = new Style { Font = { Bold = true } };
        var names = Enumerable.Range(1, 10)
            .Select(i => new string((char)('a' + i), 255))
            .ToHashSet(StringComparer.Ordinal);

        // Act
        foreach (var name in names)
        {
            spreadsheet.AddStyle(style, name, StyleNameVisibility.Visible);
        }

        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var zip = await ZipArchive.CreateAsync(stream, Token);
        using var stylesXml = await zip.GetStylesXmlStreamAsync(Token);
        await VerifyXml(stylesXml);
    }

    [Fact]
    public async Task Spreadsheet_AddStyle_NamedStyleWithInvalidVisibility()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, SpreadCheetahOptions, Token);
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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, SpreadCheetahOptions, Token);
        var style = new Style { Font = { Bold = true } };

        // Act & Assert
        Assert.ThrowsAny<ArgumentException>(() => spreadsheet.AddStyle(style, name!));
    }

    [Fact]
    public async Task Spreadsheet_AddStyle_NamedStyleWithTooLongName()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, SpreadCheetahOptions, Token);
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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, SpreadCheetahOptions, Token);

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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, SpreadCheetahOptions, Token);
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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, SpreadCheetahOptions, Token);
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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, SpreadCheetahOptions, Token);

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
