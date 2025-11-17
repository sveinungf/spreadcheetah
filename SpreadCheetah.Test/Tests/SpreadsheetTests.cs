using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using Polyfills;
using SpreadCheetah.Images;
using SpreadCheetah.Styling;
using SpreadCheetah.Test.Extensions;
using SpreadCheetah.Test.Helpers;
using SpreadCheetah.Worksheets;
using System.IO.Compression;
using Color = System.Drawing.Color;
using DataValidation = SpreadCheetah.Validations.DataValidation;
using Fill = SpreadCheetah.Styling.Fill;
using SpreadsheetAssert = SpreadCheetah.TestHelpers.Assertions.SpreadsheetAssert;
using Table = SpreadCheetah.Tables.Table;
using TableStyle = SpreadCheetah.Tables.TableStyle;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public async Task Spreadsheet_CreateNew_EmptyIsValid()
    {
        // Arrange
        using var stream = new MemoryStream();

        // Act
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Name", token: Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
    }

    [Fact]
    public async Task Spreadsheet_Finish_ValidSpreadsheetAfterDispose()
    {
        // Arrange
        using var stream = new MemoryStream();

        // Act
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await spreadsheet.StartWorksheetAsync("Name", token: Token);
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
    }

    [Theory]
    [InlineData(false)]
    public async Task Spreadsheet_CreateNew_EmptyToWriteOnlyStream(bool asyncOnly)
    {
        // Arrange
        using var stream = asyncOnly
            ? new AsyncWriteOnlyMemoryStream()
            : new WriteOnlyMemoryStream();

        // Act
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Name", token: Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        Assert.True(stream.Position > 0);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_CreateNew_WithCompressionLevel(SpreadCheetahCompressionLevel level)
    {
        // Arrange
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { CompressionLevel = level };

        // Act
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Name", token: Token);
        await spreadsheet.AddRowAsync([new("Hello"), new(123)], Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_Finish_ThrowsWhenNoWorksheet(bool hasWorksheet)
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);

        if (hasWorksheet)
            await spreadsheet.StartWorksheetAsync("Book 1", token: Token);

        // Act
        var exception = await Record.ExceptionAsync(async () => await spreadsheet.FinishAsync(Token));

        // Assert
        Assert.NotEqual(hasWorksheet, exception is SpreadCheetahException);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_Finish_OnlyNoteFilesWhenThereIsNote(bool hasNote)
    {
        // Arrange
        var options = new SpreadCheetahOptions { DefaultDateTimeFormat = null };

        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Book 1", token: Token);

        if (hasNote)
            spreadsheet.AddNote("B5", "My note");

        // Act
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var zip = await ZipArchive.CreateAsync(stream, Token);
        var filenames = zip.Entries.Select(x => x.FullName).ToList();
        Assert.Equal(hasNote, filenames.Contains("xl/comments1.xml", StringComparer.OrdinalIgnoreCase));
        Assert.Equal(hasNote, filenames.Contains("xl/drawings/vmlDrawing1.vml", StringComparer.OrdinalIgnoreCase));
        Assert.Equal(hasNote, filenames.Contains("xl/worksheets/_rels/sheet1.xml.rels", StringComparer.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_Finish_OnlyStylesXmlWhenThereIsStyling(bool hasStyle)
    {
        // Arrange
        var options = new SpreadCheetahOptions();
        if (!hasStyle)
            options.DefaultDateTimeFormat = null;

        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token))
        {
            await spreadsheet.StartWorksheetAsync("Book 1", token: Token);

            // Act
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        using var zip = await ZipArchive.CreateAsync(stream, Token);
        Assert.Equal(hasStyle, zip.GetEntry("xl/styles.xml") is not null);
    }

    [Fact]
    public async Task Spreadsheet_Finish_WorksheetWithAutoFilterAndOtherFeaturesThatAffectWorksheetXml()
    {
        // Arrange
        using var stream = new MemoryStream();
        using var pngStream = EmbeddedResources.GetStream("red-1x1.png");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
        var options = new WorksheetOptions
        {
            AutoFilter = new AutoFilterOptions("A1:F1"),
            FrozenColumns = 2,
            FrozenRows = 1,
            Visibility = WorksheetVisibility.Hidden,
            ShowGridlines = false
        };

        options.Column(2).Width = 80;

        await spreadsheet.StartWorksheetAsync("Hidden sheet", options, Token);

        var validation = DataValidation.TextLengthLessThan(50);
        spreadsheet.AddDataValidation("A2:A100", validation);
        spreadsheet.AddImage(ImageCanvas.OriginalSize("B1"), embeddedImage);
        spreadsheet.AddNote("C1", "My note");
        spreadsheet.MergeCells("B2:F3");

        await spreadsheet.StartWorksheetAsync("Visible sheet", token: Token);

        // Act
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
    }

    [Fact]
    public async Task Spreadsheet_Finish_WorksheetWithTableAndOtherFeaturesThatAffectWorksheetXml()
    {
        // Arrange
        using var stream = new MemoryStream();
        using var pngStream = EmbeddedResources.GetStream("red-1x1.png");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream, Token);
        var options = new WorksheetOptions
        {
            FrozenColumns = 2,
            FrozenRows = 1,
            Visibility = WorksheetVisibility.Hidden,
            ShowGridlines = false
        };

        options.Column(2).Width = 80;

        await spreadsheet.StartWorksheetAsync("Hidden sheet", options, Token);

        var validation = DataValidation.TextLengthLessThan(50);
        spreadsheet.AddDataValidation("A2:A100", validation);
        spreadsheet.AddImage(ImageCanvas.OriginalSize("B1"), embeddedImage);
        spreadsheet.AddNote("C1", "My note");
        spreadsheet.MergeCells("B2:F3");
        spreadsheet.StartTable(new Table(TableStyle.Light8));

        await spreadsheet.AddHeaderRowAsync(["Header"], token: Token);

        await spreadsheet.StartWorksheetAsync("Visible sheet", token: Token);

        // Act
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
    }

    [Fact]
    public async Task Spreadsheet_Dispose_DisposeTwice()
    {
        // Arrange
        var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);
#pragma warning disable
        spreadsheet.Dispose();
#pragma warning restore

        // Act
        var exception = Record.Exception(() => spreadsheet.Dispose());

        // Assert
        Assert.Null(exception);
    }

    [Fact]
    public async Task Spreadsheet_DisposeAsync_DisposeTwice()
    {
        // Arrange
        var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);
        await spreadsheet.DisposeAsync();

        // Act
        var exception = await Record.ExceptionAsync(() => spreadsheet.DisposeAsync().AsTask());

        // Assert
        Assert.Null(exception);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(10)]
    public async Task Spreadsheet_NextRowNumber_SingleWorksheet(int rowsAdded)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);

        foreach (var row in Enumerable.Range(1, rowsAdded).Select(_ => new DataCell("Value")))
        {
            await spreadsheet.AddRowAsync([row], Token);
        }

        // Act
        var result = spreadsheet.NextRowNumber;

        // Assert
        Assert.Equal(rowsAdded + 1, result);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(10)]
    public async Task Spreadsheet_NextRowNumber_MultipleWorksheets(int rowsAdded)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);

        foreach (var row in Enumerable.Range(1, 5).Select(_ => new DataCell("Value")))
        {
            await spreadsheet.AddRowAsync([row], Token);
        }

        await spreadsheet.StartWorksheetAsync("Sheet 2", token: Token);

        foreach (var row in Enumerable.Range(1, rowsAdded).Select(_ => new DataCell("Value")))
        {
            await spreadsheet.AddRowAsync([row], Token);
        }

        // Act
        var result = spreadsheet.NextRowNumber;

        // Assert
        Assert.Equal(rowsAdded + 1, result);
    }

    [Theory]
    [InlineData("OneWord")]
    [InlineData("With whitespace")]
    [InlineData("With trailing whitespace ")]
    [InlineData(" With leading whitespace")]
    [InlineData("With-Special-Characters!#¤%&")]
    [InlineData("With'Single'Quotes")]
    [InlineData("With\"Quotation\"Marks")]
    [InlineData("WithNorwegianCharactersÆØÅ")]
    [InlineData("Exactly 31 Characters Long Name")]
    public async Task Spreadsheet_StartWorksheet_ValidName(string name)
    {
        // Arrange
        using var stream = new MemoryStream();

        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            // Act
            await spreadsheet.StartWorksheetAsync(name, token: Token);
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheets = actual.WorkbookPart?.Workbook.Sheets?.Cast<Sheet>();
        Assert.NotNull(sheets);
        var sheet = Assert.Single(sheets);
        Assert.Equal(name, sheet.Name?.Value);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("This name is over the 31 character limit in Excel")]
    [InlineData("'Starting with single quote")]
    [InlineData("Ending with single quote'")]
    [InlineData("With / forward slash")]
    [InlineData("With \\ backward slash")]
    [InlineData("With ? question mark")]
    [InlineData("With * asterisk")]
    [InlineData("With [ left square bracket")]
    [InlineData("With ] right square bracket")]
    public async Task Spreadsheet_StartWorksheet_InvalidName(string name)
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);

        // Act
        var exception = await Record.ExceptionAsync(async () => await spreadsheet.StartWorksheetAsync(name, token: Token));

        // Assert
        Assert.NotNull(exception);
    }

    [Fact]
    public async Task Spreadsheet_StartWorksheet_NullName()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);

        // Act
        var exception = await Assert.ThrowsAsync<ArgumentNullException>(async () => await spreadsheet.StartWorksheetAsync(null!, token: Token));

        // Assert
        Assert.Equal("name", exception.ParamName);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Spreadsheet_StartWorksheet_ThrowsOnDuplicateName(bool duplicateName, bool sameCasing)
    {
        // Arrange
        const string name = "Sheet";
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync(name, token: Token);
        var nextName = duplicateName switch
        {
            true when sameCasing => name,
            true when !sameCasing => name.ToUpperInvariant(),
            _ => "Sheet 2"
        };

        // Act
        var exception = await Record.ExceptionAsync(async () => await spreadsheet.StartWorksheetAsync(nextName, token: Token));

        // Assert
        Assert.Equal(duplicateName, exception != null);
    }

    [Theory]
    [InlineData(2)]
    [InlineData(10)]
    [InlineData(1000)]
    [InlineData(2000)]
    public async Task Spreadsheet_StartWorksheet_MultipleWorksheets(int count)
    {
        // Arrange
        var sheetNames = Enumerable.Range(1, count).Select(x => "Sheet " + x).ToList();
        using var stream = new MemoryStream();
        var spreadsheetOptions = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, spreadsheetOptions, Token);

        // Act
        foreach (var name in sheetNames)
        {
            await spreadsheet.StartWorksheetAsync(name, token: Token);
            await spreadsheet.AddRowAsync([new DataCell(name)], Token);
        }

        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheets = SpreadsheetAssert.Sheets(stream);
        Assert.Equal(count, sheets.Count);
        Assert.Equal(sheetNames, sheets.Select(x => x.Name));
        Assert.Equal(sheetNames, sheets.Select(x => x["A1"].StringValue));
    }

    [Theory]
    [InlineData(2)]
    [InlineData(10)]
    [InlineData(100)]
    [InlineData(2000)]
    [InlineData(16383)]
    public async Task Spreadsheet_StartWorksheet_WorksheetWithMultipleColumnOptions(int count)
    {
        // Arrange
        var columnWidths = Enumerable.Range(1, count).Select(x => 20d + (x % 100)).ToList();
        using var stream = new MemoryStream();
        var spreadsheetOptions = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, spreadsheetOptions, Token);
        var options = new WorksheetOptions();

        // Act
        foreach (var (i, columnWidth) in columnWidths.Index())
        {
            options.Column(i + 1).Width = columnWidth;
        }

        await spreadsheet.StartWorksheetAsync("Sheet", options, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(columnWidths.Count, sheet.Columns.Count);
        Assert.Equal(columnWidths, sheet.Columns.Select(x => x.Width), new DoubleEqualityComparer(0.01d));
    }

    [Theory]
    [InlineData(false, WorksheetVisibility.Hidden)]
    [InlineData(false, WorksheetVisibility.VeryHidden)]
    [InlineData(true, WorksheetVisibility.Hidden)]
    [InlineData(true, WorksheetVisibility.VeryHidden)]
    public async Task Spreadsheet_StartWorksheet_Hidden(bool firstHidden, WorksheetVisibility visibility)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);

        // Act
        var options1 = new WorksheetOptions { Visibility = firstHidden ? visibility : WorksheetVisibility.Visible };
        await spreadsheet.StartWorksheetAsync("Sheet 1", options1, Token);
        var options2 = new WorksheetOptions { Visibility = !firstHidden ? visibility : WorksheetVisibility.Visible };
        await spreadsheet.StartWorksheetAsync("Sheet 2", options2, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        var openXmlVisibility = visibility.GetOpenXmlWorksheetVisibility();
        using var actual = SpreadsheetDocument.Open(stream, true);
        var workbook = actual.WorkbookPart!.Workbook;
        var sheets = workbook.Sheets!.Cast<Sheet>().ToList();
        Assert.Equal(firstHidden ? openXmlVisibility : null, sheets[0].State?.Value);
        Assert.Equal(!firstHidden ? openXmlVisibility : null, sheets[1].State?.Value);

        if (firstHidden)
        {
            var workbookView = workbook.BookViews!.GetFirstChild<WorkbookView>();
            Assert.Equal(1u, workbookView?.ActiveTab?.Value);
            Assert.Equal(1u, workbookView?.FirstSheet?.Value);
        }
    }

    [Theory]
    [InlineData(0.1)]
    [InlineData(10d)]
    [InlineData(100d)]
    [InlineData(255d)]
    [InlineData(null)]
    public async Task Spreadsheet_StartWorksheet_ColumnWidth(double? width)
    {
        // Arrange
        var worksheetOptions = new WorksheetOptions();
        worksheetOptions.Column(1).Width = width;
        const double defaultWidthInExcelUi = 8.43;
        const double defaultWidthInXml = defaultWidthInExcelUi + 0.71062;
        var expectedWidth = width ?? defaultWidthInXml;

        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);

        // Act
        await spreadsheet.StartWorksheetAsync("My sheet", worksheetOptions, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(expectedWidth, sheet.Column("A").Width, 5);
    }

    [Theory]
    [InlineData(10d)]
    [InlineData(100d)]
    [InlineData(255d)]
    [InlineData(null)]
    public async Task Spreadsheet_StartWorksheet_CustomDefaultColumnWidth(double? width)
    {
        // Arrange
        var worksheetOptions = new WorksheetOptions { DefaultColumnWidth = width };
        const double defaultWidthInExcelUi = 8.43;
        const double defaultWidthInXml = defaultWidthInExcelUi + 0.71062;
        var expectedWidth = width ?? defaultWidthInXml;

        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);

        // Act
        await spreadsheet.StartWorksheetAsync("My sheet", worksheetOptions, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(expectedWidth, sheet.Column("A").Width, 2);
    }

    [Fact]
    public async Task Spreadsheet_StartWorksheet_CustomDefaultColumnWidthOverridden()
    {
        // Arrange
        const double defaultWidth = 50;
        const double widthForThirdColumn = 40;
        var worksheetOptions = new WorksheetOptions { DefaultColumnWidth = defaultWidth };
        worksheetOptions.Column(3).Width = widthForThirdColumn;

        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);

        // Act
        await spreadsheet.StartWorksheetAsync("My sheet", worksheetOptions, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(defaultWidth, sheet.Column("A").Width, 2);
        Assert.Equal(defaultWidth, sheet.Column("B").Width, 2);
        Assert.Equal(widthForThirdColumn, sheet.Column("C").Width, 2);
        Assert.Equal(defaultWidth, sheet.Column("D").Width, 2);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_StartWorksheet_ColumnHidden(bool hidden)
    {
        // Arrange
        var worksheetOptions = new WorksheetOptions();
        var columnOptions = worksheetOptions.Column(1);
        columnOptions.Hidden = hidden;
        var expectedHidden = hidden;

        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            // Act
            await spreadsheet.StartWorksheetAsync("My sheet", worksheetOptions, Token);
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets.Single();
        var actualHidden = worksheet.Column(1).Hidden;
        Assert.Equal(expectedHidden, actualHidden);
    }

    [Theory]
    [InlineData(1, null)]
    [InlineData(3, null)]
    [InlineData(null, 1)]
    [InlineData(null, 5)]
    [InlineData(1, 2)]
    [InlineData(5, 4)]
    public async Task Spreadsheet_StartWorksheet_Freezing(int? columns, int? rows)
    {
        // Arrange
        var worksheetOptions = new WorksheetOptions
        {
            FrozenColumns = columns,
            FrozenRows = rows
        };

        var expectedColumnName = SpreadsheetUtility.GetColumnName((columns ?? 0) + 1);
        var expectedCellReference = $"{expectedColumnName}{(rows ?? 0) + 1}";
        var expectedActivePane = columns switch
        {
            not null when rows is not null => PaneValues.BottomRight,
            not null => PaneValues.TopRight,
            _ => PaneValues.BottomLeft
        };

        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            // Act
            await spreadsheet.StartWorksheetAsync("My sheet", worksheetOptions, Token);
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var worksheet = actual.WorkbookPart!.WorksheetParts.Select(x => x.Worksheet).Single();
        var sheetView = worksheet.SheetViews!.Cast<SheetView>().Single();
        Assert.Equal(PaneStateValues.Frozen, sheetView.Pane!.State!.Value);
        Assert.Equal(columns, (int?)sheetView.Pane.HorizontalSplit?.Value);
        Assert.Equal(rows, (int?)sheetView.Pane.VerticalSplit?.Value);
        Assert.Equal(expectedCellReference, sheetView.Pane.TopLeftCell?.Value);
        Assert.Equal(expectedActivePane, sheetView.Pane.ActivePane?.Value);
    }

    [Fact]
    public async Task Spreadsheet_StartWorksheet_NoFreezing()
    {
        // Arrange
        var worksheetOptions = new WorksheetOptions();
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            // Act
            await spreadsheet.StartWorksheetAsync("My sheet", worksheetOptions, Token);
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var worksheet = actual.WorkbookPart!.WorksheetParts.Select(x => x.Worksheet).Single();
        Assert.Null(worksheet.SheetViews);
    }

    [Theory]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(null, true)]
    public async Task Spreadsheet_StartWorksheet_ShowGridLines(bool? expectedShowGridLines, bool actualShowGridLines)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);

        // Act
        var sheetOptions = new WorksheetOptions { ShowGridlines = expectedShowGridLines };
        await spreadsheet.StartWorksheetAsync("Sheet 1", sheetOptions, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        Assert.Equal(worksheet.ShowGridLines, actualShowGridLines);
    }

    [Fact]
    public async Task Spreadsheet_StartWorksheet_AutoFilter()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        const string cellRange = "A1:B10";
        var autoFilter = new AutoFilterOptions(cellRange);

        // Act
        var sheetOptions = new WorksheetOptions { AutoFilter = autoFilter };
        await spreadsheet.StartWorksheetAsync("Sheet 1", sheetOptions, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        Assert.Equal(cellRange, worksheet.AutoFilter.Range.RangeAddress.ToString());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_StartWorksheet_ThrowsWhenHasFinished(bool hasFinished)
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Book 1", token: Token);

        if (hasFinished)
            await spreadsheet.FinishAsync(Token);

        // Act
        var exception = await Record.ExceptionAsync(async () => await spreadsheet.StartWorksheetAsync("Book 2", token: Token));

        // Assert
        Assert.Equal(hasFinished, exception is SpreadCheetahException);
    }

    [Fact]
    public async Task Spreadsheet_AddStyle_DuplicateStylesReturnTheSameStyleId()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Book 1", token: Token);

        var style1 = new Style { Fill = new Fill { Color = Color.Bisque } };
        var style2 = style1 with { };

        // Act
        var style1Id = spreadsheet.AddStyle(style1);
        var style2Id = spreadsheet.AddStyle(style2);

        // Assert
        Assert.Equal(style1Id.Id, style2Id.Id);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_AddStyle_DefaultStyleGettingTheExpectedStyleId(bool defaultStyle)
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Book 1", token: Token);

        var style = new Style();
        if (!defaultStyle)
            style.Format = NumberFormat.Standard(StandardNumberFormat.Fraction);

        // Act
        var styleId = spreadsheet.AddStyle(style);

        // Assert
        Assert.Equal(defaultStyle, styleId.Id == 0);
    }
}
