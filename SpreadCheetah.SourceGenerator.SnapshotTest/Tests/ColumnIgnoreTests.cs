using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class ColumnIgnoreTests
{
    [Fact]
    public Task ColumnIgnore_ClassPropertyWithAnotherSpreadCheetahAttribute()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithColumnIgnore
            {
                [{|SPCH1008:CellFormat("#.0#")|}]
                [ColumnIgnore]
                [{|SPCH1008:ColumnWidth(10)|}]
                public decimal Value { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithColumnIgnore))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync();
    }

    [Fact]
    public Task ColumnIgnore_ClassPropertyWithAnotherUnrelatedAttribute()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;
            using System.ComponentModel.DataAnnotations;

            namespace MyNamespace;

            public class ClassWithColumnIgnore
            {
                [Required]
                [ColumnIgnore]
                public decimal Value { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithColumnIgnore))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync();
    }
}
