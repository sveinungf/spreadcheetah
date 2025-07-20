using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class ContextTests
{
    [Fact]
    public Task Context_InternalAccessibility()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public record MyRecord(string? Name);
            
            [WorksheetRow(typeof(MyRecord))]
            internal partial class MyContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task Context_DefaultAccessibility()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public record MyRecord(string? Name);

            [WorksheetRow(typeof(MyRecord))]
            partial class MyContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task Context_TwoWorksheetRowAttributes()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;
            
            public record RecordA(string? Name);
            public record RecordB(int Id);

            [WorksheetRow(typeof(RecordA))]
            [WorksheetRow(typeof(RecordB))]
            public partial class MyContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task Context_TwoSimilarWorksheetRowAttributes()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace NamespaceA
            {
                public record MyRecord(string? Name);
            }

            namespace NamespaceB
            {
                public record MyRecord(string? Name);
            }

            namespace MyNamespace
            {
                [WorksheetRow(typeof(NamespaceA.MyRecord))]
                [WorksheetRow(typeof(NamespaceB.MyRecord))]
                public partial class MyContext : WorksheetRowContext;
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task Context_TwoWorksheetRowAttributesWhenTheFirstTypeEmitsWarning()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using System;
            
            namespace MyNamespace;
            
            public record RecordA(string? Name, UriBuilder? Uri);
            public record RecordB(string? Name);
            
            [WorksheetRow(typeof(RecordA))]
            [WorksheetRow(typeof(RecordB))]
            public partial class MyContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task Context_TwoContextClasses()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public record RecordA(string? Name);
            public record RecordB(int Id);

            [WorksheetRow(typeof(RecordA))]
            public partial class MyContextA : WorksheetRowContext;

            [WorksheetRow(typeof(RecordB))]
            public partial class MyContextB : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
}
