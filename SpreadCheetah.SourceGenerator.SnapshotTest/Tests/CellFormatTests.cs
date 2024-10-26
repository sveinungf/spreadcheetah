using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class CellFormatTests
{
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
}
