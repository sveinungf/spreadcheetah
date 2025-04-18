using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class ColumnOrderTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public Task ColumnOrder_ClassWithColumnOrderForAllProperties()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithColumnOrderForAllProperties
            {
                [ColumnOrder(2)]
                public string FirstName { get; set; } = "";

                [ColumnOrder(3)]
                public string? MiddleName { get; set; }

                [ColumnOrder(1)]
                public string LastName { get; set; } = "";

                [ColumnOrder(5)]
                public int Age { get; set; }

                [ColumnOrder(4)]
                public bool Employed { get; set; }

                [ColumnOrder(6)]
                public double Score { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithColumnOrderForAllProperties))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task ColumnOrder_ClassWithColumnOrderForSomeProperties()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithColumnOrderForSomeProperties
            {
                public string FirstName { get; set; } = "";

                [ColumnOrder(-1000)]
                public string? MiddleName { get; set; }

                public string LastName { get; set; } = "";

                [ColumnOrder(500)]
                public int Age { get; set; }

                public bool Employed { get; set; }

                [ColumnOrder(2)]
                public double Score { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithColumnOrderForSomeProperties))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task ColumnOrder_ClassWithDuplicateColumnOrdering()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithDuplicateColumnOrdering
            {
                [ColumnOrder(1)]
                public string PropertyA { get; set; }
                [{|SPCH1003:ColumnOrder(1)|}]
                public string PropertyB { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithDuplicateColumnOrdering))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }

    [Fact]
    public Task ColumnOrder_ClassWithDuplicateColumnOrderingAcrossInheritance()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            [InheritColumns]
            public class ClassWithDuplicateColumnOrdering : ClassWithDuplicateColumnOrderingBase
            {
                [{|SPCH1003:ColumnOrder(1)|}]
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
        return context.RunAsync(Token);
    }
}
