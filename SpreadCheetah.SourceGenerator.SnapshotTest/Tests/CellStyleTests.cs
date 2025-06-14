using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class CellStyleTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public Task CellStyle_ClassWithMultipleAttributes()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithCellStyle
            {
                [CellStyle("Italic")]
                public string? FirstName { get; set; }
            
                [CellStyle("Bold")]
                public string? LastName { get; set; }
            
                public int YearOfBirth { get; set; }
            
                [CellStyle("Bold")]
                public string? Initials { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithCellStyle))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task CellStyle_ContextWithTwoClassesWithCellStyle()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithCellStyle
            {
                [CellStyle("Italic")]
                public string? FirstName { get; set; }
            
                [CellStyle("Bold")]
                public string? LastName { get; set; }
            
                public int YearOfBirth { get; set; }
            
                [CellStyle("Bold")]
                public string? Initials { get; set; }
            }

            public class Class2WithCellStyle
            {
                [CellStyle("Red")]
                public string? Id { get; set; }
            
                [CellStyle("Bold")]
                public decimal Price { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithCellStyle))]
            [WorksheetRow(typeof(Class2WithCellStyle))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task CellStyle_ClassWithEmptyCellStyle()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;
        
            namespace MyNamespace;
        
            public class ClassWithCellStyle
            {
                [CellStyle({|SPCH1006:""|})]
                public string? FirstName { get; set; }
            }
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }

    [Fact]
    public Task CellStyle_ClassWithFormulas()
    {
        // Arrange
        const string source = """
            using SpreadCheetah;
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithCellStyle
            {
                [CellStyle("Italic")]
                public Formula FirstName { get; set; }
            
                [CellStyle("Bold")]
                public Formula? LastName { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithCellStyle))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
}
