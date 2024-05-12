using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class WorksheetRowGeneratorColumnOrderTests
{
    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithColumnOrderForAllProperties()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models.ColumnOrdering;

            namespace MyNamespace;
            
            [WorksheetRow(typeof(ClassWithColumnOrderForAllProperties))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithColumnOrderForSomeProperties()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models.ColumnOrdering;

            namespace MyNamespace;
            
            [WorksheetRow(typeof(ClassWithColumnOrderForSomeProperties))]
            public partial class MyGenRowContext : WorksheetRowContext;
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

            namespace MyNamespace;

            public class ClassWithDuplicateColumnOrdering
            {
                [ColumnOrder(1)]
                public string PropertyA { get; set; }
                [ColumnOrder(1)]
                public string PropertyB { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithDuplicateColumnOrdering))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithDuplicateColumnOrderingAcrossInheritance()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            [InheritColumns]
            public class ClassWithDuplicateColumnOrdering : ClassWithDuplicateColumnOrderingBase
            {
                [ColumnOrder(1)]
                public string PropertyA { get; set; }
            }

            public class ClassWithDuplicateColumnOrderingBase
            {
                [ColumnOrder(1)]
                public string PropertyB { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithDuplicateColumnOrdering))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
}
