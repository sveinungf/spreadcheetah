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
        using var pngStream = EmbeddedResources.GetStream("one-red-pixel.png");
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
        using var pngStream = EmbeddedResources.GetStream("one-red-pixel.png");
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
    public async Task Spreadsheet_EmbedImage_Png()
    {
        // Arrange
        using var pngStream = EmbeddedResources.GetStream("one-red-pixel.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream);

        // Act
        var result = await spreadsheet.EmbedImageAsync(pngStream);

        // Assert
        await spreadsheet.StartWorksheetAsync("Sheet 1");
        await spreadsheet.FinishAsync();

        Assert.Equal(1, result.Height);
        Assert.Equal(1, result.Width);

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
        using var pngStream = new AsyncReadOnlyMemoryStream(EmbeddedResources.GetStream("one-red-pixel.png"));
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream);

        // Act
        var exception = await Record.ExceptionAsync(() => spreadsheet.EmbedImageAsync(pngStream).AsTask());

        // Assert
        Assert.Null(exception);
    }
}
