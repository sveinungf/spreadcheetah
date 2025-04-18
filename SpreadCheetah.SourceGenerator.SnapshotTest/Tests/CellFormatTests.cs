using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class CellFormatTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public Task CellFormat_ClassWithMultipleAttributes()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.Styling;

            namespace MyNamespace;

            public class ClassWithCellFormat
            {
                [CellFormat("#.0#")]
                public decimal First { get; set; }
            
                [CellFormat(StandardNumberFormat.Exponential)]
                public decimal? Second { get; set; }
            
                public int Year { get; set; }
            
                [CellFormat("#.0#")]
                public int? Score { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithCellFormat))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task CellFormat_ClassWithInvalidFormats()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.Styling;

            namespace MyNamespace;

            public class ClassWithCellFormat
            {
                [CellFormat({|SPCH1006:"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"|})]
                public decimal First { get; set; }

                [CellFormat({|SPCH1006:(StandardNumberFormat)100|})]
                public decimal? Second { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithCellFormat))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }

    [Fact]
    public Task CellFormat_ClassPropertyWithCellFormatAndCellStyle()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithCellFormat
            {
                [CellFormat("#.0#")]
                [{|SPCH1008:CellStyle("My style")|}]
                public decimal Value { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithCellFormat))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }
}
