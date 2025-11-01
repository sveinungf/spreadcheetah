using SpreadCheetah.SourceGenerator.Test.Models;
using SpreadCheetah.SourceGenerator.Test.Models.Contexts;
using SpreadCheetah.SourceGenerator.Test.Models.InheritColumns;
using SpreadCheetah.TestHelpers.Assertions;
using SpreadCheetah.TestHelpers.Extensions;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class InheritColumnsTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Theory]
    [InlineData(ObjectType.Class)]
    [InlineData(ObjectType.RecordClass)]
    public async Task InheritColumns_ObjectWithInheritance(ObjectType type)
    {
        // Arrange
        var ctx = InheritanceContext.Default;

        using var stream = new MemoryStream();
        await using var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await s.StartWorksheetAsync("Sheet", token: Token);

        // Act
        var task = type switch
        {
            ObjectType.Class => s.AddHeaderRowAsync(ctx.ClassDog, token: Token),
            ObjectType.RecordClass => s.AddHeaderRowAsync(ctx.RecordClassDog, token: Token),
            _ => throw new NotImplementedException()
        };

        await task;
        await s.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal("CanWalk", sheet["A1"].StringValue);
        Assert.Equal("DateOfBirth", sheet["B1"].StringValue);
        Assert.Equal("Breed", sheet["C1"].StringValue);
        Assert.Equal(3, sheet.CellCount);
    }

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
        Assert.Equal(["Derived"], sheet.Row(1).Cells.StringValues());
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
        Assert.Equal(["Base", "Derived"], sheet.Row(1).Cells.StringValues());
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
        Assert.Equal(["Derived", "Base"], sheet.Row(1).Cells.StringValues());
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
        Assert.Equal(["Derived", "Base", "Leaf"], sheet.Row(1).Cells.StringValues());
    }
}