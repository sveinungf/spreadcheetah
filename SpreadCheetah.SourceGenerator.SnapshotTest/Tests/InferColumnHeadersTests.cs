using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class InferColumnHeadersTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public Task InferColumnHeaders_ClassWithValidReferences()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models.ColumnHeader;

            namespace MyNamespace;

            [InferColumnHeaders(typeof(ColumnHeaderResources), Prefix = "Header_")]
            public class ClassWithPropertyReferenceColumnHeaders
            {
                public string? FirstName { get; set; }
                public string? LastName { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithPropertyReferenceColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task InferColumnHeaders_ClassWithMissingReference()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ColumnHeaders
            {
                public static string DateOfBirth => "Date of birth";
            }

            [InferColumnHeaders(typeof(ColumnHeaders))]
            public class ClassWithInferColumnHeaders
            {
                public string? Name { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithInferColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }
}
