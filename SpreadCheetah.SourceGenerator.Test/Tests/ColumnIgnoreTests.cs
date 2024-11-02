using SpreadCheetah.SourceGenerator.Test.Models.ColumnIgnore;
using SpreadCheetah.TestHelpers.Assertions;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class ColumnIgnoreTests
{
    [Fact]
    public async Task ColumnIgnore_ClassWithMultipleProperties()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var obj = new ClassWithMultipleProperties
        {
            Id = 1,
            Name = "Foo",
            Price = 199.90m
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj, ColumnIgnoreContext.Default.ClassWithMultipleProperties);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(obj.Name, sheet["A1"].StringValue);
        Assert.Single(sheet.Columns);
    }
}
