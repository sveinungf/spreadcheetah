using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class CellValueTruncateTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public Task CellValueTruncate_ClassWithCellValueTruncate()
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
    public Task CellValueTruncate_ClassWithMultipleCellValueTruncate()
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
    public Task CellValueTruncate_ClassWithCellValueTruncateWithInvalidLength()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithCellValueTruncate
            {
                [CellValueTruncate({|SPCH1006:0|})]
                public string? Name { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithCellValueTruncate))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }

    [Fact]
    public Task CellValueTruncate_ClassWithCellValueTruncateOnInvalidType()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithCellValueTruncate
            {
                [{|SPCH1005:CellValueTruncate(10)|}]
                public int Year { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithCellValueTruncate))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }
}
