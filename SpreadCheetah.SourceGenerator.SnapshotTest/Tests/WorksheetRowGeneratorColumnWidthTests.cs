using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class WorksheetRowGeneratorColumnWidthTests
{
    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithColumnWidth()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithColumnWidth
            {
                [ColumnWidth(10)]
                public string? Name { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithColumnWidth))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithMultipleColumnWidths()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithMultipleColumnWidths
            {
                [ColumnWidth(20d)]
                public string? FirstName { get; set; }

                [ColumnWidth(25.67)]
                public string? LastName { get; set; }

                public int YearOfBirth { get; set; }

                [ColumnWidth(0.9999)]
                public string? Initials { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithMultipleColumnWidths))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_ContextWithTwoClassesWithColumnWidth()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class Class1WithColumnWidth
            {
                [ColumnWidth(10)]
                public string? Name { get; set; }
            }

            public class Class2WithColumnWidth
            {
                [ColumnWidth(20)]
                public int? Age { get; set; }
            }
            
            [WorksheetRow(typeof(Class1WithColumnWidth))]
            [WorksheetRow(typeof(Class2WithColumnWidth))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithColumnWidthWithInvalidWidth()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithColumnWidth
            {
                [ColumnWidth(300)]
                public string? Name { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithColumnWidth))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source, onlyDiagnostics: true);
    }
}
