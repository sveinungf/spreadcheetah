using SpreadCheetah.Tables;
using SpreadCheetah.Test.Helpers;
using SpreadCheetah.TestHelpers.Extensions;
using SpreadsheetAssert = SpreadCheetah.TestHelpers.Assertions.SpreadsheetAssert;

namespace SpreadCheetah.Test.Tests;

public class HeaderRowTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Theory, CombinatorialData]
    public async Task Spreadsheet_AddHeaderRow_Success(RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        string[] headerNames = ["ID", "First name", "Last name", "Age"];

        // Act
        await spreadsheet.AddHeaderRowAsync(headerNames, rowType);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(headerNames, sheet.Row(1).StringValues());
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_AddHeaderRow_WithNullValue(RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        string[] headerNames = ["ID", null!, "Last name", "Age"];

        // Act
        await spreadsheet.AddHeaderRowAsync(headerNames, rowType);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(headerNames, sheet.Row(1).StringValues());
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_AddHeaderRow_Empty(RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        const string subsequentRowValue = "Hello";

        // Act
        await spreadsheet.AddHeaderRowAsync([], rowType);
        await spreadsheet.AddRowAsync([new DataCell(subsequentRowValue)], rowType);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Empty(sheet.Row(1));
        Assert.Equal(subsequentRowValue, sheet["A2"].StringValue);
    }

    [Theory]
    [InlineData(RowCollectionType.Array)]
    [InlineData(RowCollectionType.List)]
    public async Task Spreadsheet_AddHeaderRow_Null(RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        // Act
        var exception = await Record.ExceptionAsync(rowType switch
        {
            RowCollectionType.Array => async () => await spreadsheet.AddHeaderRowAsync(null!, token: Token),
            RowCollectionType.List => async () => await spreadsheet.AddHeaderRowAsync((null as List<string>)!, token: Token),
            _ => throw new ArgumentOutOfRangeException(nameof(rowType), rowType, null)
        });

        // Assert
        Assert.IsType<ArgumentNullException>(exception);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_AddHeaderRow_HeaderLength(
        [CombinatorialValues(255, 256)] int length, bool activeTable)
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var headerName = new string('a', length);

        if (activeTable)
            spreadsheet.StartTable(new Table(TableStyle.Light1));

        // Act
        var exception = await Record.ExceptionAsync(() => spreadsheet.AddHeaderRowAsync([headerName], token: Token).AsTask());

        // Assert
        Assert.Equal(activeTable && length > 255, exception is not null);
    }
}
