using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using SpreadCheetah.Styling;
using SpreadCheetah.Test.Helpers;
using SpreadCheetah.Worksheets;
using System.IO.Compression;
using Xunit;
using Color = System.Drawing.Color;
using Fill = SpreadCheetah.Styling.Fill;

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
        Assert.Equal(hasWorksheet, exception == null);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_Finish_OnlyStylesXmlWhenThereIsStyling(bool hasStyle)
    {
        // Arrange
        var options = new SpreadCheetahOptions();
        if (!hasStyle)
            options.DefaultDateTimeNumberFormat = null;

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

    [Theory]
    [InlineData("OneWord")]
    [InlineData("With whitespace")]
    [InlineData("With trailing whitespace ")]
    [InlineData(" With leading whitespace")]
    [InlineData("With-Special-Characters!#¤%&")]
    [InlineData("With'Single'Quotes")]
    [InlineData("With\"Quotation\"Marks")]
    [InlineData("WithNorwegianCharactersÆØÅ")]
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
    [InlineData(null)]
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
    public async Task Spreadsheet_StartWorksheet_MultipleWorksheets(int count)
    {
        // Arrange
        var sheetNames = Enumerable.Range(1, count).Select(x => "Sheet " + x).ToList();
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            // Act
            foreach (var name in sheetNames)
            {
                await spreadsheet.StartWorksheetAsync(name);
                await spreadsheet.AddRowAsync(new DataCell(name));
            }

            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheets = actual.WorkbookPart!.Workbook.Sheets!.Cast<Sheet>().ToList();
        var worksheets = actual.WorkbookPart.WorksheetParts.Select(x => x.Worksheet);
        var cells = worksheets.Select(x => x.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single());
        Assert.Equal(count, sheets.Count);
        Assert.Equal(sheetNames, sheets.Select(x => x.Name.Value));
        Assert.Equal(sheetNames, cells.Select(x => x.InnerText));
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
        SpreadsheetAssert.Valid(stream);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets.Single();
        var actualWidth = worksheet.Column(1).Width;
        Assert.Equal(expectedWidth, actualWidth, 5);
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

        var expectedColumnName = CellReferenceHelper.GetExcelColumnName((columns ?? 0) + 1);
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
        Assert.Equal(!hasFinished, exception is null);
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
            style.NumberFormat = NumberFormats.Fraction;

        // Act
        var styleId = spreadsheet.AddStyle(style);

        // Assert
        Assert.Equal(defaultStyle, styleId.Id == 0);
    }
}
