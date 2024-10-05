using SpreadCheetah.SourceGenerator.Test.Models.InheritColumns;
using SpreadCheetah.TestHelpers.Assertions;
using SpreadCheetah.TestHelpers.Extensions;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class InheritColumnsTests
{
    [Fact]
    public async Task InheritColumns_DerivedClassWithoutInheritColumns()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var obj = new DerivedClassWithoutInheritColumns
        {
            BaseClassProperty = "Base",
            DerivedClassProperty = "Derived"
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj, InheritColumnsContext.Default.DerivedClassWithoutInheritColumns);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(["Derived"], sheet.Row(1).StringValues());
    }

    [Fact]
    public async Task InheritColumns_ClassWithInheritedColumnsFirst()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var obj = new ClassWithInheritedColumnsFirst
        {
            BaseClassProperty = "Base",
            DerivedClassProperty = "Derived"
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj, InheritColumnsContext.Default.ClassWithInheritedColumnsFirst);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(["Base", "Derived"], sheet.Row(1).StringValues());
    }

    [Fact]
    public async Task InheritColumns_ClassWithInheritedColumnsLast()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var obj = new ClassWithInheritedColumnsLast
        {
            BaseClassProperty = "Base",
            DerivedClassProperty = "Derived"
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj, InheritColumnsContext.Default.ClassWithInheritedColumnsLast);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(["Derived", "Base"], sheet.Row(1).StringValues());
    }
}