using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class AccessibilityTests
{
    [Fact]
    public Task Accessibility_ContextClassWithInternalAccessibility()
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
    public Task Accessibility_ContextClassWithDefaultAccessibility()
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
}
