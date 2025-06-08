using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using SpreadCheetah.Images;
using SpreadCheetah.Test.Helpers;
using SpreadCheetah.Worksheets;
using System.IO.Compression;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetImageTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public async Task Spreadsheet_EmbedImage_NullStream()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);

        // Act
        var exception = await Record.ExceptionAsync(() => spreadsheet.EmbedImageAsync(null!, Token).AsTask());

        // Assert
        var concreteException = Assert.IsType<ArgumentNullException>(exception);
        Assert.Equal("stream", concreteException.ParamName);
    }

    [Fact]
    public async Task Spreadsheet_EmbedImage_WriteOnlyStream()
    {
        // Arrange
        using var writeOnlyStream = new WriteOnlyMemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);

        // Act
        var exception = await Record.ExceptionAsync(() => spreadsheet.EmbedImageAsync(writeOnlyStream, Token).AsTask());

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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        pngStream.Position = pngStream.Length;

        // Act
        var exception = await Record.ExceptionAsync(() => spreadsheet.EmbedImageAsync(pngStream, Token).AsTask());

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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);

        // Act
        var exception = await Record.ExceptionAsync(() => spreadsheet.EmbedImageAsync(pngStream, Token).AsTask());

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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        await spreadsheet.FinishAsync(Token);

        // Act
        var exception = await Record.ExceptionAsync(() => spreadsheet.EmbedImageAsync(pngStream, Token).AsTask());

        // Assert
        var concreteException = Assert.IsType<SpreadCheetahException>(exception);
        Assert.Contains(nameof(Spreadsheet.FinishAsync), concreteException.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Spreadsheet_EmbedImage_InvalidFile()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);
        var bytes = "Invalid file"u8.ToArray();
        var imageStream = new MemoryStream(bytes);

        // Act
        var exception = await Record.ExceptionAsync(() => spreadsheet.EmbedImageAsync(imageStream, Token).AsTask());

        // Assert
        var concreteException = Assert.IsType<ArgumentException>(exception);
        Assert.Equal("stream", concreteException.ParamName);
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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);

        // Act
        var result = await spreadsheet.EmbedImageAsync(pngStream, Token);

        // Assert
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        await spreadsheet.FinishAsync(Token);
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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);

        // Act
        var exception = await Record.ExceptionAsync(() => spreadsheet.EmbedImageAsync(pngStream, Token).AsTask());

        // Assert
        Assert.Null(exception);
    }

    [Fact]
    public async Task Spreadsheet_EmbedImage_PngWithInvalidResolution()
    {
        // Arrange
        using var pngStream = EmbeddedResources.GetStream("red-1x1.invalidpng");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);

        // Act
        var exception = await Record.ExceptionAsync(() => spreadsheet.EmbedImageAsync(pngStream, Token).AsTask());

        // Assert
        var concreteException = Assert.IsType<ArgumentOutOfRangeException>(exception);
        Assert.Contains("width", concreteException.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Spreadsheet_AddImage_Png()
    {
        // Arrange
        const string reference = "B3";
        using var pngStream = EmbeddedResources.GetStream("green-266x183.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        var canvas = ImageCanvas.OriginalSize(reference.AsSpan());

        // Act
        spreadsheet.AddImage(canvas, embeddedImage);

        // Assert
        await spreadsheet.FinishAsync(Token);
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
    public async Task Spreadsheet_AddImage_PngWithTransparentBackground()
    {
        // Arrange
        const string reference = "B3";
        using var pngStream = EmbeddedResources.GetStream("yellow-500x500-transparent.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        var canvas = ImageCanvas.OriginalSize(reference.AsSpan());

        // Act
        spreadsheet.AddImage(canvas, embeddedImage);

        // Assert
        await spreadsheet.FinishAsync(Token);
        SpreadsheetAssert.Valid(outputStream);

        using var workbook = new XLWorkbook(outputStream);
        var worksheet = Assert.Single(workbook.Worksheets);
        var picture = Assert.Single(worksheet.Pictures);
        Assert.Equal(reference, picture.TopLeftCell.Address.ToString());
        Assert.Equal(XLPictureFormat.Png, picture.Format);
        Assert.Equal(500, picture.OriginalWidth);
        Assert.Equal(500, picture.OriginalHeight);
        Assert.Equal(0, picture.Top);
        Assert.Equal(0, picture.Left);
        Assert.Equal("Image 1", picture.Name);
    }

    [Fact]
    public async Task Spreadsheet_AddImage_PngWithLargeResolution()
    {
        // Arrange
        const string reference = "B3";
        using var pngStream = EmbeddedResources.GetStream("orange-10000x10000.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        var canvas = ImageCanvas.OriginalSize(reference.AsSpan());

        // Act
        spreadsheet.AddImage(canvas, embeddedImage);

        // Assert
        await spreadsheet.FinishAsync(Token);
        SpreadsheetAssert.Valid(outputStream);

        using var workbook = new XLWorkbook(outputStream);
        var worksheet = workbook.Worksheets.Single();
        var picture = Assert.Single(worksheet.Pictures);
        Assert.Equal(reference, picture.TopLeftCell.Address.ToString());
        Assert.Equal(XLPictureFormat.Png, picture.Format);
        Assert.Equal(10000, picture.OriginalWidth);
        Assert.Equal(10000, picture.OriginalHeight);
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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        await spreadsheet.StartWorksheetAsync(worksheetName, token: Token);
        var canvas = ImageCanvas.OriginalSize(reference.AsSpan());

        // Act
        spreadsheet.AddImage(canvas, embeddedImage);

        // Assert
        await spreadsheet.FinishAsync(Token);
        SpreadsheetAssert.Valid(outputStream);

        using var workbook = new XLWorkbook(outputStream);
        var worksheet1 = Assert.Single(workbook.Worksheets, x => !string.Equals(worksheetName, x.Name, StringComparison.Ordinal));
        Assert.Empty(worksheet1.Pictures);

        var worksheet2 = Assert.Single(workbook.Worksheets, x => string.Equals(worksheetName, x.Name, StringComparison.Ordinal));
        var picture = Assert.Single(worksheet2.Pictures);
        Assert.Equal(reference, picture.TopLeftCell.Address.ToString());
        Assert.Equal(XLPictureFormat.Png, picture.Format);
        Assert.Equal(266, picture.OriginalWidth);
        Assert.Equal(183, picture.OriginalHeight);
        Assert.Equal(0, picture.Top);
        Assert.Equal(0, picture.Left);
        Assert.Equal("Image 1", picture.Name);
    }

    [Theory]
    [InlineData(2)]
    [InlineData(10)]
    [InlineData(1000)]
    public async Task Spreadsheet_AddImage_MultiplePng(int count)
    {
        // Arrange
        using var pngStream = EmbeddedResources.GetStream("red-1x1.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImages = new List<EmbeddedImage>(count);
        for (var i = 0; i < count; i++)
        {
            pngStream.Position = 0;
            var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
            embeddedImages.Add(embeddedImage);
        }

        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        var row = 1;

        // Act
        foreach (var embeddedImage in embeddedImages)
        {
            var reference = "A" + row;
            var canvas = ImageCanvas.OriginalSize(reference.AsSpan());
            spreadsheet.AddImage(canvas, embeddedImage);
            row++;
        }

        // Assert
        await spreadsheet.FinishAsync(Token);
        SpreadsheetAssert.Valid(outputStream);

        using var workbook = new XLWorkbook(outputStream);
        var worksheet = workbook.Worksheets.Single();
        var pictures = worksheet.Pictures.ToList();
        Assert.Equal(count, pictures.Count);
        Assert.All(pictures, x => Assert.Equal(1, x.OriginalWidth));
        Assert.All(pictures, x => Assert.Equal(1, x.OriginalHeight));
        Assert.Distinct(pictures.Select(x => x.Name), StringComparer.Ordinal);
        Assert.Distinct(pictures.Select(x => x.TopLeftCell.Address.ToString()), StringComparer.Ordinal);
    }

    [Fact]
    public async Task Spreadsheet_AddImage_SamePngUsedInMultipleCells()
    {
        // Arrange
        var references = new[] { "A10", "B3", "J2" };
        using var pngStream = EmbeddedResources.GetStream("green-266x183.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);

        // Act
        foreach (var reference in references)
        {
            var canvas = ImageCanvas.OriginalSize(reference.AsSpan());
            spreadsheet.AddImage(canvas, embeddedImage);
        }

        // Assert
        await spreadsheet.FinishAsync(Token);
        SpreadsheetAssert.Valid(outputStream);

        using var workbook = new XLWorkbook(outputStream);
        var worksheet = workbook.Worksheets.Single();
        var pictures = worksheet.Pictures.ToList();
        Assert.Equal(references.Length, pictures.Count);
        Assert.All(pictures, x => Assert.Equal(266, x.OriginalWidth));
        Assert.All(pictures, x => Assert.Equal(183, x.OriginalHeight));
        Assert.Equal(references, pictures.Select(x => x.TopLeftCell.Address.ToString()));
        Assert.Distinct(pictures.Select(x => x.Name), StringComparer.Ordinal);
    }

    [Fact]
    public async Task Spreadsheet_AddImage_SamePngUsedInMultipleWorksheets()
    {
        // Arrange
        var worksheet1References = new[] { "A10", "B3", "J2" };
        var worksheet2References = new[] { "B10", "C1", "H20" };
        using var pngStream = EmbeddedResources.GetStream("red-1x1.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);

        // Act
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        foreach (var reference in worksheet1References)
        {
            var canvas = ImageCanvas.OriginalSize(reference.AsSpan());
            spreadsheet.AddImage(canvas, embeddedImage);
        }

        await spreadsheet.StartWorksheetAsync("Sheet 2", token: Token);
        foreach (var reference in worksheet2References)
        {
            var canvas = ImageCanvas.OriginalSize(reference.AsSpan());
            spreadsheet.AddImage(canvas, embeddedImage);
        }

        // Assert
        await spreadsheet.FinishAsync(Token);
        SpreadsheetAssert.Valid(outputStream);

        using var workbook = new XLWorkbook(outputStream);
        var worksheet1 = workbook.Worksheets.First();
        var worksheet1Pictures = worksheet1.Pictures.ToList();
        Assert.Equal(worksheet1References, worksheet1Pictures.Select(x => x.TopLeftCell.Address.ToString()));

        var worksheet2 = workbook.Worksheets.Skip(1).First();
        var worksheet2Pictures = worksheet2.Pictures.ToList();
        Assert.Equal(worksheet2References, worksheet2Pictures.Select(x => x.TopLeftCell.Address.ToString()));

        Assert.Distinct(worksheet1Pictures.Concat(worksheet2Pictures).Select(x => x.Name), StringComparer.Ordinal);
    }

    [Fact]
    public async Task Spreadsheet_AddImage_SamePngUsedInManyWorksheets()
    {
        // Arrange
        const int worksheetCount = 100;
        const string reference = "A1";
        using var pngStream = EmbeddedResources.GetStream("red-1x1.png");
        using var outputStream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, options, Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);

        // Act
        for (var i = 0; i < worksheetCount; i++)
        {
            await spreadsheet.StartWorksheetAsync($"Sheet {i}", token: Token);
            var canvas = ImageCanvas.OriginalSize(reference.AsSpan());
            spreadsheet.AddImage(canvas, embeddedImage);
        }

        // Assert
        await spreadsheet.FinishAsync(Token);
        SpreadsheetAssert.Valid(outputStream);

        using var workbook = new XLWorkbook(outputStream);
        var actualWorksheets = workbook.Worksheets.ToList();
        Assert.All(actualWorksheets, x => Assert.Single(x.Pictures));
        Assert.All(actualWorksheets, x => Assert.Equal(reference, x.Pictures.Single().TopLeftCell.Address.ToString()));
    }

    [Fact]
    public async Task Spreadsheet_AddImage_MultiplePngUsedInMultipleWorksheets()
    {
        // Arrange
        var imageFilenames = new[] { "blue-250x250.png", "green-266x183.png", "red-1x1.png", "yellow-1x1.png" };
        var worksheet1References = new[] { "A10", "B3" };
        var worksheet2References = new[] { "B10", "C1" };

        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImages = new Queue<EmbeddedImage>();

        foreach (var filename in imageFilenames)
        {
            using var pngStream = EmbeddedResources.GetStream("red-1x1.png");
            var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
            embeddedImages.Enqueue(embeddedImage);
        }

        // Act
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        foreach (var reference in worksheet1References)
        {
            var canvas = ImageCanvas.OriginalSize(reference.AsSpan());
            spreadsheet.AddImage(canvas, embeddedImages.Dequeue());
        }

        await spreadsheet.StartWorksheetAsync("Sheet 2", token: Token);
        foreach (var reference in worksheet2References)
        {
            var canvas = ImageCanvas.OriginalSize(reference.AsSpan());
            spreadsheet.AddImage(canvas, embeddedImages.Dequeue());
        }

        // Assert
        await spreadsheet.FinishAsync(Token);
        SpreadsheetAssert.Valid(outputStream);

        using var workbook = new XLWorkbook(outputStream);
        var worksheet1 = workbook.Worksheets.First();
        var worksheet1Pictures = worksheet1.Pictures.ToList();
        Assert.Equal(worksheet1References, worksheet1Pictures.Select(x => x.TopLeftCell.Address.ToString()));

        var worksheet2 = workbook.Worksheets.Skip(1).First();
        var worksheet2Pictures = worksheet2.Pictures.ToList();
        Assert.Equal(worksheet2References, worksheet2Pictures.Select(x => x.TopLeftCell.Address.ToString()));

        Assert.Distinct(worksheet1Pictures.Concat(worksheet2Pictures).Select(x => x.Name), StringComparer.Ordinal);
    }

    [Fact]
    public async Task Spreadsheet_AddImage_EmbeddedImageFromOtherSpreadsheet()
    {
        // Arrange
        using var pngStream = EmbeddedResources.GetStream("green-266x183.png");
        using var stream = new MemoryStream();
        using var otherStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await using var otherSpreadsheet = await Spreadsheet.CreateNewAsync(otherStream, cancellationToken: Token);
        var embeddedImage = await otherSpreadsheet.EmbedImageAsync(pngStream, Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        var canvas = ImageCanvas.OriginalSize("A1".AsSpan());

        // Act
        var exception = Record.Exception(() => spreadsheet.AddImage(canvas, embeddedImage));

        // Assert
        var concreteException = Assert.IsType<SpreadCheetahException>(exception);
        Assert.Contains("spreadsheet", concreteException.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(true, true, XLPicturePlacement.MoveAndSize)]
    [InlineData(true, false, XLPicturePlacement.Move)]
    [InlineData(false, false, XLPicturePlacement.FreeFloating)]
    public async Task Spreadsheet_AddImage_Anchor(bool moveWithCells, bool resizeWithCells, XLPicturePlacement expectedPlacement)
    {
        // Arrange
        using var pngStream = EmbeddedResources.GetStream("red-1x1.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        var canvas = ImageCanvas.FillCell("A1".AsSpan(), moveWithCells, resizeWithCells);

        // Act
        spreadsheet.AddImage(canvas, embeddedImage);

        // Assert
        await spreadsheet.FinishAsync(Token);
        SpreadsheetAssert.Valid(outputStream);

        using var workbook = new XLWorkbook(outputStream);
        var worksheet = workbook.Worksheets.Single();
        var picture = Assert.Single(worksheet.Pictures);
        Assert.Equal(expectedPlacement, picture.Placement);
    }

    [Theory]
    [InlineData(1, 1)]
    [InlineData(10, 20)]
    [InlineData(266, 183)]
    [InlineData(1000, 1000)]
    public async Task Spreadsheet_AddImage_PngWithCustomDimension(int width, int height)
    {
        // Arrange
        using var pngStream = EmbeddedResources.GetStream("green-266x183.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        var canvas = ImageCanvas.Dimensions("D4".AsSpan(), width, height);

        // Act
        spreadsheet.AddImage(canvas, embeddedImage);

        // Assert
        await spreadsheet.FinishAsync(Token);
        SpreadsheetAssert.Valid(outputStream);
        using var zip = new ZipArchive(outputStream);
        using var drawingXml = zip.GetDrawingXmlStream();
        await VerifyXml(drawingXml);
    }

    [Theory]
    [InlineData(0.1)]
    [InlineData(0.5)]
    [InlineData(1.0)]
    [InlineData(5.46)]
    public async Task Spreadsheet_AddImage_PngWithCustomScale(float scale)
    {
        // Arrange
        using var pngStream = EmbeddedResources.GetStream("green-266x183.png");
        using var outputStream = File.Create("test.xlsx");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        var canvas = ImageCanvas.Scaled("D4".AsSpan(), scale);

        // Act
        spreadsheet.AddImage(canvas, embeddedImage);

        // Assert
        await spreadsheet.FinishAsync(Token);
        SpreadsheetAssert.Valid(outputStream);
        using var zip = new ZipArchive(outputStream);
        using var drawingXml = zip.GetDrawingXmlStream();
        await VerifyXml(drawingXml);
    }

    [Theory]
    [InlineData(0.002)]
    [InlineData(376.0)]
    public async Task Spreadsheet_AddImage_PngWithInvalidCustomScale(float scale)
    {
        // Arrange
        using var pngStream = EmbeddedResources.GetStream("green-266x183.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        var canvas = ImageCanvas.Scaled("D4".AsSpan(), scale);

        // Act & Assert
        Assert.ThrowsAny<ArgumentOutOfRangeException>(() => spreadsheet.AddImage(canvas, embeddedImage));
    }

    [Fact]
    public async Task Spreadsheet_AddImage_PngWithFillCell()
    {
        // Arrange
        using var pngStream = EmbeddedResources.GetStream("green-266x183.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        var canvas = ImageCanvas.FillCell("B3".AsSpan());

        // Act
        spreadsheet.AddImage(canvas, embeddedImage);

        // Assert
        await spreadsheet.FinishAsync(Token);
        SpreadsheetAssert.Valid(outputStream);

        using var package = new ExcelPackage(outputStream);
        var worksheet = Assert.Single(package.Workbook.Worksheets);
        var drawing = Assert.Single(worksheet.Drawings);
        Assert.Equal("C4", drawing.To.ToCellReferenceString());
    }

    [Theory]
    [InlineData("C4")]
    [InlineData("C20")]
    [InlineData("D7")]
    [InlineData("P4")]
    public async Task Spreadsheet_AddImage_PngWithFillCellRange(string lowerRightReference)
    {
        // Arrange
        using var pngStream = EmbeddedResources.GetStream("green-266x183.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        var canvas = ImageCanvas.FillCells("B3".AsSpan(), lowerRightReference.AsSpan());

        // Act
        spreadsheet.AddImage(canvas, embeddedImage);

        // Assert
        await spreadsheet.FinishAsync(Token);
        SpreadsheetAssert.Valid(outputStream);

        using var package = new ExcelPackage(outputStream);
        var worksheet = Assert.Single(package.Workbook.Worksheets);
        var drawing = Assert.Single(worksheet.Drawings);
        Assert.Equal(lowerRightReference, drawing.To.ToCellReferenceString());
    }

    [Theory]
    [InlineData(-100, -80)]
    [InlineData(-2, 5)]
    [InlineData(0, 0)]
    [InlineData(3, -8)]
    public async Task Spreadsheet_AddImage_PngWithOffset(int left, int top)
    {
        // Arrange
        const string reference = "T20";
        using var pngStream = EmbeddedResources.GetStream("green-266x183.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        var canvas = ImageCanvas.OriginalSize(reference.AsSpan());
        var options = new ImageOptions { Offset = ImageOffset.Pixels(left, top, 0, 0) };

        // Act
        spreadsheet.AddImage(canvas, embeddedImage, options);

        // Assert
        await spreadsheet.FinishAsync(Token);
        SpreadsheetAssert.Valid(outputStream);
        using var zip = new ZipArchive(outputStream);
        using var drawingXml = zip.GetDrawingXmlStream();
        await VerifyXml(drawingXml);
    }

    [Theory]
    [InlineData(-10, -10, 10, 10)]
    [InlineData(0, 0, 0, 0)]
    [InlineData(0, 0, 0, 10)]
    [InlineData(0, 0, 10, 0)]
    [InlineData(0, 10, 0, 0)]
    [InlineData(10, 0, 0, 0)]
    [InlineData(10, 10, 10, 10)]
    public async Task Spreadsheet_AddImage_PngWithFillCellsAndOffsets(int left, int top, int right, int bottom)
    {
        // Arrange
        const string reference = "C3";
        using var pngStream = EmbeddedResources.GetStream("green-266x183.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        var canvas = ImageCanvas.FillCell(reference.AsSpan());
        var options = new ImageOptions { Offset = ImageOffset.Pixels(left, top, right, bottom) };

        // Act
        spreadsheet.AddImage(canvas, embeddedImage, options);

        // Assert
        await spreadsheet.FinishAsync(Token);
        SpreadsheetAssert.Valid(outputStream);
        using var zip = new ZipArchive(outputStream);
        using var drawingXml = zip.GetDrawingXmlStream();
        await VerifyXml(drawingXml);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_AddImage_PngWithDefaultCanvas(bool withWorksheetOptions)
    {
        // Arrange
        using var pngStream = EmbeddedResources.GetStream("green-266x183.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
        var options = withWorksheetOptions ? new WorksheetOptions() : null;
        await spreadsheet.StartWorksheetAsync("Sheet 1", options, Token);
        ImageCanvas canvas = default;

        // Act
        spreadsheet.AddImage(canvas, embeddedImage);

        // Assert
        await spreadsheet.FinishAsync(Token);
        SpreadsheetAssert.Valid(outputStream);
        using var zip = new ZipArchive(outputStream);
        using var drawingXml = zip.GetDrawingXmlStream();
        var verifySettings = new VerifySettings();
        verifySettings.IgnoreParametersForVerified();
        await VerifyXml(drawingXml, verifySettings);
    }

    [Fact]
    public async Task Spreadsheet_AddImage_CustomColumnWidths()
    {
        // Arrange
        using var pngStream = EmbeddedResources.GetStream("green-266x183.png");
        using var outputStream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(outputStream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
        var options = new WorksheetOptions();
        options.Column(3).Width = 20.0;
        options.Column(5).Width = 5.0;
        options.Column(6).Width = 5.0;
        options.Column(7).Width = 5.0;
        options.Column(8).Width = 30.0;
        options.Column(10).Width = 15.0;
        options.Column(11).Width = 15.0;

        await spreadsheet.StartWorksheetAsync("Sheet 1", options, Token);
        var canvas = ImageCanvas.Dimensions("B2".AsSpan(), 1000, 1000);

        // Act
        spreadsheet.AddImage(canvas, embeddedImage);

        // Assert
        await spreadsheet.FinishAsync(Token);
        SpreadsheetAssert.Valid(outputStream);
        using var zip = new ZipArchive(outputStream);
        using var drawingXml = zip.GetDrawingXmlStream();
        await VerifyXml(drawingXml);
    }
}
