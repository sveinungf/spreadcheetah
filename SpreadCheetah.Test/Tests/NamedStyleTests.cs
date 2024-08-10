using OfficeOpenXml;
using SpreadCheetah.Styling;
using SpreadCheetah.Test.Helpers;

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

    // TODO: Test for DateTime stuff
}
