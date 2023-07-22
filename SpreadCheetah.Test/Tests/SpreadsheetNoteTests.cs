using ClosedXML.Excel;
using SpreadCheetah.Test.Helpers;
using Xunit;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetNoteTests
{
    [Fact]
    public async Task Spreadsheet_AddNote_Success()
    {
        // Arrange
        const string noteText = "My note";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        // Act
        spreadsheet.AddNote("A1", noteText);
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.True(actualCell.HasComment);
        var actualNote = actualCell.GetComment();
        Assert.Equal(noteText, actualNote.Text);
    }
}
