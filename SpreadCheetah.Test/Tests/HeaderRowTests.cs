using SpreadCheetah.Test.Helpers;
using SpreadCheetah.TestHelpers.Extensions;
using SpreadsheetAssert = SpreadCheetah.TestHelpers.Assertions.SpreadsheetAssert;

namespace SpreadCheetah.Test.Tests;

public class HeaderRowTests
{
    [Theory, CombinatorialData]
    public async Task Spreadsheet_AddHeaderRow_Success(RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        string?[] headerNames = ["ID", "First name", "Last name", "Age"];

        // Act
        await spreadsheet.AddHeaderRowAsync(headerNames, rowType);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(headerNames, sheet.Row(1).StringValues());
    }

    // TODO: With 'null' value
    // TODO: With empty collection
    // TODO: With null collection
}
