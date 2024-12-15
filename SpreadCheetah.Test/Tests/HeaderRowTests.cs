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

    [Theory, CombinatorialData]
    public async Task Spreadsheet_AddHeaderRow_WithNullValue(RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        string?[] headerNames = ["ID", null, "Last name", "Age"];

        // Act
        await spreadsheet.AddHeaderRowAsync(headerNames, rowType);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(headerNames, sheet.Row(1).StringValues());
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_AddHeaderRow_Empty(RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        // Act
        await spreadsheet.AddHeaderRowAsync([], rowType);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Empty(sheet.Row(1));
    }

    [Theory]
    [InlineData(RowCollectionType.Array)]
    [InlineData(RowCollectionType.List)]
    public async Task Spreadsheet_AddHeaderRow_Null(RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        // Act
        var task = rowType switch
        {
            RowCollectionType.Array => spreadsheet.AddHeaderRowAsync(null!),
            RowCollectionType.List => spreadsheet.AddHeaderRowAsync((null as List<string?>)!),
            _ => throw new ArgumentOutOfRangeException(nameof(rowType), rowType, null)
        };

        // Assert
        await Assert.ThrowsAsync<ArgumentNullException>(task.AsTask);
    }
}
