using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using SpreadCheetah.Test.Helpers;
using Xunit;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetNoteTests
{
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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        await spreadsheet.AddRowAsync(new Cell(cellValue));

        // Act
        spreadsheet.AddNote(reference, noteText);
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(reference);
        Assert.Equal(cellValue, actualCell.Value);
        Assert.True(actualCell.HasComment);
        var actualNote = actualCell.GetComment();
        Assert.Equal(noteText, actualNote.Text);
        var p = actualNote.Position;
        Assert.Equal(expectedPosition, new[] { p.Column, p.ColumnOffset, p.Row, p.RowOffset });
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
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        // Act
        spreadsheet.AddNote(reference, noteText);
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var vmlDrawing = sheetPart.VmlDrawingParts.Single();
        using var vmlStream = vmlDrawing.GetStream();
        using var reader = new StreamReader(vmlStream);
        var vml = await reader.ReadToEndAsync();
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
}
