using SpreadCheetah.Test.Helpers;
using Xunit;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetNoteTests
{
    [Fact]
    public async Task Spreadsheet_AddNote_Success()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        // Act
        spreadsheet.AddNote("A1", "My note");
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        // TODO: Verify
    }
}
