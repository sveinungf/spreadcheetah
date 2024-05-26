using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class WorksheetRowGeneratorCellValueTruncateTests
{
    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithCellValueTruncate()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithCellValueTruncate
            {
                [CellValueTruncate(10)]
                public string? Name { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithCellValueTruncate))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithMultipleCellValueTruncate()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithMultipleCellValueTruncate
            {
                [CellValueTruncate(20)]
                public string? FirstName { get; set; }

                [CellValueTruncate(20)]
                public string? LastName { get; set; }

                public int YearOfBirth { get; set; }

                [CellValueTruncate(5)]
                public string? Initials { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithMultipleCellValueTruncate))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithCellValueTruncateWithInvalidLength()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithCellValueTruncate
            {
                [CellValueTruncate(0)]
                public string? Name { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithCellValueTruncate))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithCellValueTruncateOnInvalidType()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithCellValueTruncate
            {
                [CellValueTruncate(10)]
                public int Year { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithCellValueTruncate))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
}
