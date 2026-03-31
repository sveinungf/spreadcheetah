using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class DefaultColumnWidthTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public Task DefaultColumnWidth_ClassWithDefaultColumnWidth()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            [DefaultColumnWidth(10)]
            public class ClassWithDefaultColumnWidth
            {
                public int Id { get; set; }

                [ColumnWidth(20)]
                public string? Name { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithDefaultColumnWidth))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task DefaultColumnWidth_ClassWithInvalidColumnWidth()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            [DefaultColumnWidth({|SPCH1006:300|})]
            public class ClassWithDefaultColumnWidth
            {
                public string? Name { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithDefaultColumnWidth))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }

    [Fact]
    public Task DefaultColumnWidth_BaseAndDerivedClassWithDefaultColumnWidth()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            [DefaultColumnWidth(30)]
            public class BaseClass
            {
                public int Id { get; set; }
            }

            [{|SPCH1012:DefaultColumnWidth(20)|}]
            [InheritColumns]
            public class DerivedClass : BaseClass
            {
                public string? Name { get; set; }
            }
            
            [WorksheetRow(typeof(DerivedClass))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }
}
