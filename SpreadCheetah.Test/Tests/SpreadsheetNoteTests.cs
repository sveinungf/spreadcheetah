using ClosedXML.Excel;
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
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        // Act
        spreadsheet.AddNote(reference, noteText);
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(reference);
        Assert.True(actualCell.HasComment);
        var actualNote = actualCell.GetComment();
        Assert.Equal(noteText, actualNote.Text);
        var p = actualNote.Position;
        Assert.Equal(expectedPosition, new[] { p.Column, p.ColumnOffset, p.Row, p.RowOffset });
    }
}
