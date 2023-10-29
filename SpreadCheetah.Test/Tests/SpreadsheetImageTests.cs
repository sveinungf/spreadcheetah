using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using SpreadCheetah.Test.Helpers;
using System.IO.Compression;
using Xunit;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetImageTests
{
    [Fact]
    public async Task Spreadsheet_EmbedImage_NullStream()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);

        // Act
        var exception = await Record.ExceptionAsync(() => spreadsheet.EmbedImageAsync(null!).AsTask());

        // Assert
        var concreteException = Assert.IsType<ArgumentNullException>(exception);
        Assert.Equal("stream", concreteException.ParamName);
    }

    [Fact]
    public async Task Spreadsheet_EmbedImage_WriteOnlyStream()
    {
        // Arrange
        using var writeOnlyStream = new WriteOnlyMemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);

        // Act
        var exception = await Record.ExceptionAsync(() => spreadsheet.EmbedImageAsync(writeOnlyStream).AsTask());

        // Assert
        var concreteException = Assert.IsType<ArgumentException>(exception);
        Assert.Equal("stream", concreteException.ParamName);
    }

    [Fact]
    public async Task Spreadsheet_EmbedImage_StreamAtEnd()
    {
        // Arrange
        using var pngStream = EmbeddedResources.GetStream("red-1x1.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream);
        pngStream.Position = pngStream.Length;

        // Act
        var exception = await Record.ExceptionAsync(() => spreadsheet.EmbedImageAsync(pngStream).AsTask());

        // Assert
        var concreteException = Assert.IsType<ArgumentException>(exception);
        Assert.Contains("Position", concreteException.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Spreadsheet_EmbedImage_ActiveWorksheet()
    {
        // Arrange
        using var pngStream = EmbeddedResources.GetStream("red-1x1.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream);
        await spreadsheet.StartWorksheetAsync("Sheet 1");

        // Act
        var exception = await Record.ExceptionAsync(() => spreadsheet.EmbedImageAsync(pngStream).AsTask());

        // Assert
        var concreteException = Assert.IsType<SpreadCheetahException>(exception);
        Assert.Contains("embed", concreteException.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Spreadsheet_EmbedImage_FinishedSpreadsheet()
    {
        // Arrange
        using var pngStream = EmbeddedResources.GetStream("red-1x1.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream);
        await spreadsheet.StartWorksheetAsync("Sheet 1");
        await spreadsheet.FinishAsync();

        // Act
        var exception = await Record.ExceptionAsync(() => spreadsheet.EmbedImageAsync(pngStream).AsTask());

        // Assert
        var concreteException = Assert.IsType<SpreadCheetahException>(exception);
        Assert.Contains(nameof(Spreadsheet.FinishAsync), concreteException.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("red-1x1.png", 1, 1)]
    [InlineData("green-266x183.png", 266, 183)]
    [InlineData("blue-250x250.png", 250, 250)]
    public async Task Spreadsheet_EmbedImage_Png(string filename, int expectedWidth, int expectedHeight)
    {
        // Arrange
        using var pngStream = EmbeddedResources.GetStream(filename);
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream);

        // Act
        var result = await spreadsheet.EmbedImageAsync(pngStream);

        // Assert
        await spreadsheet.StartWorksheetAsync("Sheet 1");
        await spreadsheet.FinishAsync();
        SpreadsheetAssert.Valid(outputStream);

        Assert.Equal(expectedWidth, result.Width);
        Assert.Equal(expectedHeight, result.Height);

        using var archive = new ZipArchive(outputStream);
        var entry = archive.GetEntry("xl/media/image1.png");
        Assert.NotNull(entry);

        using var actualPngStream = entry.Open();
        var actualPngBytes = await actualPngStream.ToByteArrayAsync();
        var expectedPngBytes = await pngStream.ToByteArrayAsync();
        Assert.Equal(expectedPngBytes, actualPngBytes);
    }

    [Fact]
    public async Task Spreadsheet_EmbedImage_PngFromReadOnlyStream()
    {
        // Arrange
        using var pngStream = new AsyncReadOnlyMemoryStream(EmbeddedResources.GetStream("red-1x1.png"));
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream);

        // Act
        var exception = await Record.ExceptionAsync(() => spreadsheet.EmbedImageAsync(pngStream).AsTask());

        // Assert
        Assert.Null(exception);
    }

    [Fact]
    public async Task Spreadsheet_AddImage_Png()
    {
        // Arrange
        const string reference = "B3";
        using var pngStream = EmbeddedResources.GetStream("green-266x183.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream);
        await spreadsheet.StartWorksheetAsync("Sheet 1");

        // Act
        spreadsheet.AddImage(reference, embeddedImage);

        // Assert
        await spreadsheet.FinishAsync();
        SpreadsheetAssert.Valid(outputStream);

        using var workbook = new XLWorkbook(outputStream);
        var worksheet = workbook.Worksheets.Single();
        var picture = Assert.Single(worksheet.Pictures);
        Assert.Equal(reference, picture.TopLeftCell.Address.ToString());
        Assert.Equal(XLPictureFormat.Png, picture.Format);
        Assert.Equal(266, picture.OriginalWidth);
        Assert.Equal(183, picture.OriginalHeight);
        Assert.Equal(0, picture.Top);
        Assert.Equal(0, picture.Left);
        Assert.Equal("Image 1", picture.Name);
    }

    [Fact]
    public async Task Spreadsheet_AddImage_PngInSecondWorksheet()
    {
        // Arrange
        const string reference = "A1";
        const string worksheetName = "Sheet 2";
        using var pngStream = EmbeddedResources.GetStream("green-266x183.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream);
        await spreadsheet.StartWorksheetAsync("Sheet 1");
        await spreadsheet.StartWorksheetAsync(worksheetName);

        // Act
        spreadsheet.AddImage(reference, embeddedImage);

        // Assert
        await spreadsheet.FinishAsync();
        SpreadsheetAssert.Valid(outputStream);

        using var workbook = new XLWorkbook(outputStream);
        var worksheet1 = Assert.Single(workbook.Worksheets.Where(x => !string.Equals(worksheetName, x.Name, StringComparison.Ordinal)));
        Assert.Empty(worksheet1.Pictures);

        var worksheet2 = Assert.Single(workbook.Worksheets.Where(x => string.Equals(worksheetName, x.Name, StringComparison.Ordinal)));
        var picture = Assert.Single(worksheet2.Pictures);
        Assert.Equal(reference, picture.TopLeftCell.Address.ToString());
        Assert.Equal(XLPictureFormat.Png, picture.Format);
        Assert.Equal(266, picture.OriginalWidth);
        Assert.Equal(183, picture.OriginalHeight);
        Assert.Equal(0, picture.Top);
        Assert.Equal(0, picture.Left);
        Assert.Equal("Image 1", picture.Name);
    }
}
