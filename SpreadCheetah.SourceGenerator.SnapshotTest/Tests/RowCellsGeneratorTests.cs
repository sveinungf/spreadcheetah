using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

[UsesVerify]
public class RowCellsGeneratorTests
{
    [Fact]
    public Task RowCellsGenerator_Generate_ClassWithSingleProperty()
    {
        // Arrange
        var source = @"
using SpreadCheetah.SourceGenerator.SnapshotTest.Models;

public static class TestClass
{
    public static async Task Run(Spreadsheet spreadsheet)
    {
        var obj = new ClassWithSingleProperty();
        await spreadsheet.AddAsRowAsync(obj);
    }
}";

        // Act & Assert
        return TestHelper.CompileAndVerify(source);
    }
}
