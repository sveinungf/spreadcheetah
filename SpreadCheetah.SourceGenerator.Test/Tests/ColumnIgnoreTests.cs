using SpreadCheetah.SourceGenerator.Test.Models.ColumnIgnore;
using SpreadCheetah.TestHelpers.Assertions;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class ColumnIgnoreTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public async Task ColumnIgnore_ClassWithMultipleProperties()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var obj = new ClassWithMultipleProperties
        {
            Id = 1,
            Name = "Foo",
            Price = 199.90m
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj, ColumnIgnoreContext.Default.ClassWithMultipleProperties, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(obj.Name, sheet["A1"].StringValue);
        Assert.Single(sheet.Columns);
    }
}
