using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadCheetah.Test.Helpers;
using System.IO.Compression;
using System.Text;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetNoteTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Theory]
    [InlineData("A1", new[] { 2, 1.6, 1, 1 })]
    [InlineData("A2", new[] { 2, 1.6, 1, 11 })]
    [InlineData("B1", new[] { 3, 1.6, 1, 1 })]
    [InlineData("C4", new[] { 4, 1.6, 3, 11 })]
    public async Task Spreadsheet_AddNote_Success(string reference, double[] expectedPosition)
    {
        // Arrange
        const string noteText = "My note";
        const string cellValue = "My value";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        await spreadsheet.AddRowAsync([new Cell(cellValue)], Token);

        // Act
        spreadsheet.AddNote(reference, noteText);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var valueCell = worksheet.Cell(1, 1);
        Assert.Equal(cellValue, valueCell.Value);
        var noteCell = worksheet.Cell(reference);
        Assert.True(noteCell.HasComment);
        var actualNote = noteCell.GetComment();
        Assert.Equal(noteText, actualNote.Text);
        var p = actualNote.Position;
        double[] actualPosition = [p.Column, p.ColumnOffset, p.Row, p.RowOffset];
        Assert.Equal(expectedPosition, actualPosition);
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
    public async Task Spreadsheet_AddNote_CorrectText(string noteText)
    {
        // Arrange
        const string reference = "A2";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        // Act
        spreadsheet.AddNote(reference, noteText);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var noteCell = worksheet.Cell(reference);
        Assert.Equal(noteText, noteCell.GetCellNoteText());
    }

    [Fact]
    public async Task Spreadsheet_AddNote_InTwoWorksheets()
    {
        // Arrange
        const string reference = "A2";
        const string sheet1Note = "Note 1";
        const string sheet2Note = "Note 2";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);

        // Act
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        spreadsheet.AddNote(reference, sheet1Note);
        await spreadsheet.StartWorksheetAsync("Sheet 2", token: Token);
        spreadsheet.AddNote(reference, sheet2Note);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet1 = workbook.Worksheets.First();
        var note1Cell = worksheet1.Cell(reference);
        Assert.Equal(sheet1Note, note1Cell.GetCellNoteText());
        var worksheet2 = workbook.Worksheets.Skip(1).Single();
        var note2Cell = worksheet2.Cell(reference);
        Assert.Equal(sheet2Note, note2Cell.GetCellNoteText());
    }

    [Fact]
    public async Task Spreadsheet_AddNote_TextLongerThanBufferSize()
    {
        // Arrange
        var sb = new StringBuilder();
        var counter = 1;
        while (sb.Length < SpreadCheetahOptions.MinimumBufferSize + 100)
        {
            sb.Append(counter++).Append(' ');
        }

        var noteText = sb.ToString();

        const string reference = "A2";
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        // Act
        spreadsheet.AddNote(reference, noteText);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var noteCell = worksheet.Cell(reference);
        Assert.Equal(noteText, noteCell.GetCellNoteText());
    }

    [Fact]
    public async Task Spreadsheet_AddNote_MultipleWithTextLongerThanBufferSize()
    {
        // Arrange
        var sb = new StringBuilder();
        var counter = 1;
        while (sb.Length < SpreadCheetahOptions.MinimumBufferSize + 100)
        {
            sb.Append(counter++).Append(' ');
        }

        var noteText = sb.ToString();

        var references = new[] { "A2", "B4", "C6" };
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        // Act
        foreach (var reference in references)
        {
            spreadsheet.AddNote(reference, noteText);
        }

        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var noteCells = references.Select(x => worksheet.Cell(x));
        Assert.All(noteCells, x => Assert.Equal(noteText, x.GetCellNoteText()));
    }

    [Theory]
    [InlineData(32768, false)]
    [InlineData(32769, true)]
    public async Task Spreadsheet_AddNote_ThrowsOnTooLongText(int textLength, bool exceptionExpected)
    {
        // Arrange
        var noteText = new string('a', textLength);
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        // Act
        var exception = Record.Exception(() => spreadsheet.AddNote("A1", noteText));

        // Assert
        Assert.Equal(exceptionExpected, exception is ArgumentException);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("A")]
    [InlineData("A0")]
    [InlineData(" A1")]
    [InlineData("A 1")]
    [InlineData("A1 ")]
    [InlineData("A1:A2")]
    [InlineData("$A$1")]
    public async Task Spreadsheet_AddNote_InvalidReference(string? reference)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        // Act & Assert
        Assert.ThrowsAny<ArgumentException>(() => spreadsheet.AddNote(reference!, "My note"));
    }

    [Theory]
    [InlineData("A1", "1.2pt")]
    [InlineData("A2", "8.4pt")]
    [InlineData("B1", "1.2pt")]
    [InlineData("C4", "37.2pt")]
    [InlineData("A6", "66.0pt")]
    public async Task Spreadsheet_AddNote_Margins(string reference, string expectedTopMargin)
    {
        // Arrange
        const string noteText = "My note";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        // Act
        spreadsheet.AddNote(reference, noteText);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var vmlDrawing = sheetPart.VmlDrawingParts.Single();
        using var vmlStream = vmlDrawing.GetStream();
        using var reader = new StreamReader(vmlStream);
        var vml = await reader.ReadToEndAsync(Token);
        const string styleAttribute = " style=\"";
        var styleStartIndex = vml.IndexOf(styleAttribute, StringComparison.OrdinalIgnoreCase);
        var styleStartIndexActual = styleStartIndex + styleAttribute.Length;
        var styleEndIndex = vml.IndexOf('"', styleStartIndexActual);
        var style = vml.Substring(styleStartIndexActual, styleEndIndex - styleStartIndexActual);
        var strings = style.Split(';');
        var styles = strings.Select(x => x.Split(':')).ToDictionary(x => x[0], x => x[1], StringComparer.OrdinalIgnoreCase);
        Assert.Equal("57pt", styles["margin-left"]);
        Assert.Equal(expectedTopMargin, styles["margin-top"]);
    }

    [Fact]
    public async Task Spreadsheet_AddNote_SheetXmlHasLegacyDrawing()
    {
        // Arrange
        const string noteText = "My note";
        const string cellValue = "My value";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        await spreadsheet.AddRowAsync([new Cell(cellValue)], Token);

        // Act
        spreadsheet.AddNote("A1", noteText);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var worksheet = sheetPart.Worksheet;
        var legacyDrawing = worksheet.ChildElements.OfType<LegacyDrawing>().Single();
        Assert.Equal("rId1", legacyDrawing.Id);
    }

    [Theory]
    [InlineData(3)]
    [InlineData(100)]
    [InlineData(10000)]
    public async Task Spreadsheet_AddNote_MultipleNotesInWorksheet(int count)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var notes = Enumerable.Range(1, count)
            .Select(x => SpreadsheetUtility.GetColumnName(x) + x)
            .Select(x => (Reference: x, NoteText: "Note for " + x))
            .ToList();

        // Act
        foreach (var (reference, noteText) in notes)
        {
            spreadsheet.AddNote(reference, noteText);
        }

        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var noteCells = notes.Select(x => worksheet.Cell(x.Reference));
        Assert.Equal(notes.Select(x => x.NoteText), noteCells.Select(x => x.GetCellNoteText()));
    }

    [Fact]
    public async Task Spreadsheet_AddNote_NoteInSecondWorksheet()
    {
        // Arrange
        const string reference = "A1";
        const string noteText = "My note";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        await spreadsheet.StartWorksheetAsync("Sheet 2", token: Token);

        // Act
        spreadsheet.AddNote(reference, noteText);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheets = workbook.Worksheets.ToList();
        Assert.Equal(2, worksheets.Count);
        var firstSheetCell = worksheets[0].Cell(reference);
        Assert.False(firstSheetCell.HasComment);
        var secondSheetCell = worksheets[1].Cell(reference);
        Assert.Equal(noteText, secondSheetCell.GetCellNoteText());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Spreadsheet_AddNote_ExpectedNoteFileNames(bool noteInSecondSheet)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        if (noteInSecondSheet)
            await spreadsheet.StartWorksheetAsync("Sheet 2", token: Token);

        // Act
        spreadsheet.AddNote("A1", "My note");
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var zip = new ZipArchive(stream);
        var filenames = zip.Entries.Select(x => x.FullName).ToList();
        Assert.Contains("xl/comments1.xml", filenames);
        Assert.Contains("xl/drawings/vmlDrawing1.vml", filenames);
        Assert.Contains($"xl/worksheets/_rels/sheet{(noteInSecondSheet ? 2 : 1)}.xml.rels", filenames);
        Assert.DoesNotContain($"xl/worksheets/_rels/sheet{(noteInSecondSheet ? 1 : 2)}.xml.rels", filenames);
    }

    [Theory]
    [InlineData(3)]
    [InlineData(10)]
    [InlineData(1000)]
    public async Task Spreadsheet_AddNote_NotesInMultipleWorksheets(int count)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        var notes = Enumerable.Range(1, count)
            .Select(x => SpreadsheetUtility.GetColumnName(x) + x)
            .Select(x => (Reference: x, NoteText: "Note for " + x))
            .ToList();

        // Act
        for (var i = 0; i < notes.Count; i++)
        {
            var (reference, noteText) = notes[i];
            await spreadsheet.StartWorksheetAsync("Sheet " + i, token: Token);
            spreadsheet.AddNote(reference, noteText);
        }

        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheets = workbook.Worksheets.ToList();
        Assert.Equal(notes.Count, worksheets.Count);
        var actualNoteText = Enumerable.Range(0, notes.Count)
            .Select(x => worksheets[x].Cell(notes[x].Reference))
            .Select(x => x.GetCellNoteText());

        Assert.Equal(notes.Select(x => x.NoteText), actualNoteText);
    }
}
