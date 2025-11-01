using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class InheritColumnsTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public Task InheritColumns_BaseClassWithAttribute()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            [InheritColumns]
            public class BaseClass
            {
                public string? Name { get; set; }
            }
            
            [WorksheetRow(typeof(BaseClass))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }

    [Fact]
    public Task InheritColumns_InheritedColumnsLast()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            
            namespace MyNamespace;
            
            public class BaseClass
            {
                public int BaseClassProperty { get; set; }
            }
            
            [InheritColumns(DefaultColumnOrder = InheritedColumnsOrder.InheritedColumnsLast)]
            public class DerivedClass : BaseClass
            {
                public int DerivedClassProperty { get; set; }
            }
            
            [WorksheetRow(typeof(DerivedClass))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
}
