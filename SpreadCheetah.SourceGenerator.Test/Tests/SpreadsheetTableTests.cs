using SpreadCheetah.SourceGenerator.Test.Models.Tables;
using SpreadCheetah.Tables;
using SpreadCheetah.TestHelpers.Assertions;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class SpreadsheetTableTests
{
    [Fact]
    public async Task Spreadsheet_Table_RowsFromSourceGeneratedCode()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var table = new Table(TableStyle.Light1);
        var ctx = PersonContext.Default.Person;
        var person = new Person
        {
            FirstName = "Henry",
            LastName = "Ford",
            Age = 58
        };

        // Act
        spreadsheet.StartTable(table);
        await spreadsheet.AddHeaderRowAsync(ctx);
        await spreadsheet.AddAsRowAsync(person, ctx);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        var actualTable = Assert.Single(sheet.Tables);
        Assert.Equal(["Last name", "First name", "Age"], actualTable.Columns.Select(x => x.Name));
        Assert.Equal("A1:C2", actualTable.CellRangeReference);
    }
}
