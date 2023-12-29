using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

[UsesVerify]
public class WorksheetRowGeneratorColumnOrderTests
{
    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithColumnOrderForAllProperties()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models.ColumnOrdering;
            using System;

            namespace MyNamespace;
            
            [WorksheetRow(typeof(ClassWithColumnOrderForAllProperties))]
            public partial class MyGenRowContext : WorksheetRowContext
            {
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithDuplicateColumnOrdering()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using System;

            namespace MyNamespace;

            public class ClassWithDuplicateColumnOrdering
            {
                [ColumnOrder(1)]
                public string PropertyA { get; set; }
                [ColumnOrder(1)]
                public string PropertyB { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithDuplicateColumnOrdering))]
            public partial class MyGenRowContext : WorksheetRowContext
            {
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
}
