using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using SpreadCheetah.Images;
using SpreadCheetah.Styling;
using SpreadCheetah.Test.Helpers;
using SpreadCheetah.Worksheets;
using System.IO.Compression;
using Color = System.Drawing.Color;
using DataValidation = SpreadCheetah.Validations.DataValidation;
using Fill = SpreadCheetah.Styling.Fill;
using SpreadsheetAssert = SpreadCheetah.TestHelpers.Assertions.SpreadsheetAssert;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetTests
{
    [Fact]
    public async Task Spreadsheet_CreateNew_EmptyIsValid()
    {
        // Arrange
        using var stream = new MemoryStream();

        // Act
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Name");
            await spreadsheet.FinishAsync();
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
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Name");
            await spreadsheet.FinishAsync();
        }

        // Assert
        Assert.True(stream.Position > 0);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_Finish_ThrowsWhenNoWorksheet(bool hasWorksheet)
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);

        if (hasWorksheet)
            await spreadsheet.StartWorksheetAsync("Book 1");

        // Act
        var exception = await Record.ExceptionAsync(async () => await spreadsheet.FinishAsync());

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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        await spreadsheet.StartWorksheetAsync("Book 1");

        if (hasNote)
            spreadsheet.AddNote("B5", "My note");

        // Act
        await spreadsheet.FinishAsync();

        // Assert
        using var zip = new ZipArchive(stream);
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
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options))
        {
            await spreadsheet.StartWorksheetAsync("Book 1");

            // Act
            await spreadsheet.FinishAsync();
        }

        // Assert
        using var zip = new ZipArchive(stream);
        Assert.Equal(hasStyle, zip.GetEntry("xl/styles.xml") is not null);
    }

    [Fact]
    public async Task Spreadsheet_Finish_ValidSpreadsheetBeforeDispose()
    {
        // Arrange
        using var stream = new MemoryStream();

        // Act
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Name");
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
    }

    [Fact]
    public async Task Spreadsheet_Finish_WorksheetWithAllFeaturesThatAffectWorksheetXml()
    {
        // Arrange
        using var stream = new MemoryStream();
        using var pngStream = EmbeddedResources.GetStream("red-1x1.png");
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        var embeddedImage = await spreadsheet.EmbedImageAsync(pngStream);
        var options = new WorksheetOptions
        {
            AutoFilter = new AutoFilterOptions("A1:F1"),
            FrozenColumns = 2,
            FrozenRows = 1,
            Visibility = WorksheetVisibility.Hidden
        };

        options.Column(2).Width = 80;

        await spreadsheet.StartWorksheetAsync("Sheet", options);

        var validation = DataValidation.TextLengthLessThan(50);
        spreadsheet.AddDataValidation("A2:A100", validation);
        spreadsheet.AddImage(ImageCanvas.OriginalSize("B1".AsSpan()), embeddedImage);
        spreadsheet.AddNote("C1", "My note");
        spreadsheet.MergeCells("B2:F3");

        // Act
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(10)]
    public async Task Spreadsheet_NextRowNumber_SingleWorksheet(int rowsAdded)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet 1");

        foreach (var row in Enumerable.Range(1, rowsAdded).Select(_ => new DataCell("Value")))
        {
            await spreadsheet.AddRowAsync(row);
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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet 1");

        foreach (var row in Enumerable.Range(1, 5).Select(_ => new DataCell("Value")))
        {
            await spreadsheet.AddRowAsync(row);
        }

        await spreadsheet.StartWorksheetAsync("Sheet 2");

        foreach (var row in Enumerable.Range(1, rowsAdded).Select(_ => new DataCell("Value")))
        {
            await spreadsheet.AddRowAsync(row);
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
    public async Task Spreadsheet_StartWorksheet_CorrectName(string name)
    {
        // Arrange
        using var stream = new MemoryStream();

        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            // Act
            await spreadsheet.StartWorksheetAsync(name);
            await spreadsheet.FinishAsync();
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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);

        // Act
        var exception = await Record.ExceptionAsync(async () => await spreadsheet.StartWorksheetAsync(name));

        // Assert
        Assert.NotNull(exception);
    }

    [Fact]
    public async Task Spreadsheet_StartWorksheet_NullName()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);

        // Act
        var exception = await Assert.ThrowsAsync<ArgumentNullException>(async () => await spreadsheet.StartWorksheetAsync(null!));

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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync(name);
        var nextName = duplicateName switch
        {
            true when sameCasing => name,
            true when !sameCasing => name.ToUpperInvariant(),
            _ => "Sheet 2"
        };

        // Act
        var exception = await Record.ExceptionAsync(async () => await spreadsheet.StartWorksheetAsync(nextName));

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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, spreadsheetOptions);

        // Act
        foreach (var name in sheetNames)
        {
            await spreadsheet.StartWorksheetAsync(name);
            await spreadsheet.AddRowAsync(new DataCell(name));
        }

        await spreadsheet.FinishAsync();

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
    [InlineData(16384)]
    public async Task Spreadsheet_StartWorksheet_WorksheetWithMultipleColumnOptions(int count)
    {
        // Arrange
        var columnWidths = Enumerable.Range(1, count).Select(x => 20d + (x % 100)).ToList();
        using var stream = new MemoryStream();
        var spreadsheetOptions = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, spreadsheetOptions);
        var options = new WorksheetOptions();

        // Act
        foreach (var (i, columnWidth) in columnWidths.Index())
        {
            options.Column(i + 1).Width = columnWidth;
        }

        await spreadsheet.StartWorksheetAsync("Sheet", options);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);

        // Act
        await spreadsheet.StartWorksheetAsync("Sheet 1", new WorksheetOptions { Visibility = firstHidden ? visibility : WorksheetVisibility.Visible });
        await spreadsheet.StartWorksheetAsync("Sheet 2", new WorksheetOptions { Visibility = !firstHidden ? visibility : WorksheetVisibility.Visible });
        await spreadsheet.FinishAsync();

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
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            // Act
            await spreadsheet.StartWorksheetAsync("My sheet", worksheetOptions);
            await spreadsheet.FinishAsync();
        }

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(expectedWidth, sheet.Column("A").Width, 5);
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
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            // Act
            await spreadsheet.StartWorksheetAsync("My sheet", worksheetOptions);
            await spreadsheet.FinishAsync();
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
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            // Act
            await spreadsheet.StartWorksheetAsync("My sheet", worksheetOptions);
            await spreadsheet.FinishAsync();
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
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            // Act
            await spreadsheet.StartWorksheetAsync("My sheet", worksheetOptions);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var worksheet = actual.WorkbookPart!.WorksheetParts.Select(x => x.Worksheet).Single();
        Assert.Null(worksheet.SheetViews);
    }

    [Fact]
    public async Task Spreadsheet_StartWorksheet_AutoFilter()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        const string cellRange = "A1:B10";
        var autoFilter = new AutoFilterOptions(cellRange);

        // Act
        await spreadsheet.StartWorksheetAsync("Sheet 1", new WorksheetOptions { AutoFilter = autoFilter });
        await spreadsheet.FinishAsync();

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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Book 1");

        if (hasFinished)
            await spreadsheet.FinishAsync();

        // Act
        var exception = await Record.ExceptionAsync(async () => await spreadsheet.StartWorksheetAsync("Book 2"));

        // Assert
        Assert.Equal(hasFinished, exception is SpreadCheetahException);
    }

    [Fact]
    public async Task Spreadsheet_AddStyle_DuplicateStylesReturnTheSameStyleId()
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Book 1");

        var style1 = new Style { Fill = new Fill { Color = Color.Bisque } };
        var style2 = new Style { Fill = new Fill { Color = Color.Bisque } };

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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
        await spreadsheet.StartWorksheetAsync("Book 1");

        var style = new Style();
        if (!defaultStyle)
            style.Format = NumberFormat.Standard(StandardNumberFormat.Fraction);

        // Act
        var styleId = spreadsheet.AddStyle(style);

        // Assert
        Assert.Equal(defaultStyle, styleId.Id == 0);
    }
}
