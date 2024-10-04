using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class RowTypeTests
{
    [Fact]
    public Task RowType_ClassWithNoProperties()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithNoProperties;
            
            [WorksheetRow({|SPCH1001:typeof(ClassWithNoProperties)|})]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync();
    }

    [Fact]
    public Task RowType_ClassWithNoPropertiesAndWarningsSuppressed()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithNoProperties;
            
            [WorksheetRow(typeof(ClassWithNoProperties))]
            [WorksheetRowGenerationOptions(SuppressWarnings = true)]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync();
    }

    [Fact]
    public Task RowType_ClassWithUnsupportedProperty()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;
            using System;

            namespace MyNamespace;

            public record RecordWithUnsupportedProperty(Uri HomepageUri);
            
            [WorksheetRow({|SPCH1002:typeof(RecordWithUnsupportedProperty)|})]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync();
    }

    [Fact]
    public Task RowType_ClassWithUnsupportedPropertyAndWarningsSuppressed()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;
            using System;

            namespace MyNamespace;

            public record RecordWithUnsupportedProperty(Uri HomepageUri);
            
            [WorksheetRow(typeof(RecordWithUnsupportedProperty))]
            [WorksheetRowGenerationOptions(SuppressWarnings = true)]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync();
    }
}
