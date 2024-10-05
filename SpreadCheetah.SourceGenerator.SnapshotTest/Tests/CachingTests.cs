using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class CachingTests
{
    [Fact]
    public Task Caching_IncrementalSourceGeneratorCachingCorrectly()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using System;

            namespace MyNamespace;

            public record MyRecord(string? Name);
            
            [WorksheetRow(typeof(MyRecord))]
            public partial class MyContext : WorksheetRowContext;
            """;

        // Act
        var (diagnostics, output) = TestHelper.GetGeneratedTrees<WorksheetRowGenerator>(source, ["Transform"]);

        // Assert
        Assert.Empty(diagnostics);
        var outputSource = Assert.Single(output);

        var settings = new VerifySettings();
        settings.UseDirectory("../Snapshots");
        settings.UseTypeName("G");
        return Verify(outputSource, settings);
    }
}
