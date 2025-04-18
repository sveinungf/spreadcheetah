using SpreadCheetah.SourceGenerator.Test.Models.InheritColumns;
using SpreadCheetah.TestHelpers.Assertions;
using SpreadCheetah.TestHelpers.Extensions;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class InheritColumnsTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public async Task InheritColumns_DerivedClassWithoutInheritColumns()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var obj = new DerivedClassWithoutInheritColumns
        {
            BaseClassProperty = "Base",
            DerivedClassProperty = "Derived"
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj, InheritColumnsContext.Default.DerivedClassWithoutInheritColumns, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(["Derived"], sheet.Row(1).StringValues());
    }

    [Fact]
    public async Task InheritColumns_InheritedColumnsFirst()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var obj = new ClassWithInheritedColumnsFirst
        {
            BaseClassProperty = "Base",
            DerivedClassProperty = "Derived"
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj, InheritColumnsContext.Default.ClassWithInheritedColumnsFirst, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(["Base", "Derived"], sheet.Row(1).StringValues());
    }

    [Fact]
    public async Task InheritColumns_InheritedColumnsLast()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var obj = new ClassWithInheritedColumnsLast
        {
            BaseClassProperty = "Base",
            DerivedClassProperty = "Derived"
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj, InheritColumnsContext.Default.ClassWithInheritedColumnsLast, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(["Derived", "Base"], sheet.Row(1).StringValues());
    }

    [Fact]
    public async Task InheritColumns_TwoLevelInheritance()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var obj = new TwoLevelInheritanceClassWithInheritedColumns
        {
            BaseClassProperty = "Base",
            DerivedClassProperty = "Derived",
            LeafClassProperty = "Leaf"
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj, InheritColumnsContext.Default.TwoLevelInheritanceClassWithInheritedColumns, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(["Derived", "Base", "Leaf"], sheet.Row(1).StringValues());
    }
}