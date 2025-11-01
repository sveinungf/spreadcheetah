using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

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

            namespace MyNamespace;

            public class ColumnHeaders
            {
                public static string FirstName => "First name";
                public static string LastName => "Last name";
            }

            [InferColumnHeaders(typeof(ColumnHeaders))]
            public class ClassWithInferColumnHeaders
            {
                public string? FirstName { get; set; }
                public string? LastName { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithInferColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task InferColumnHeaders_DerivedClassWithValidReferences()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ColumnHeaders
            {
                public static string FirstName => "First name";
                public static string LastName => "Last name";
            }

            [InferColumnHeaders(typeof(ColumnHeaders))]
            public abstract class BaseClassWithInferColumnHeaders
            {
            }

            public class DerivedClassWithInferColumnHeaders : BaseClassWithInferColumnHeaders
            {
                public string? FirstName { get; set; }
                public string? LastName { get; set; }
            }
            
            [WorksheetRow(typeof(DerivedClassWithInferColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task InferColumnHeaders_ClassWithValidReferencesWithPrefix()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models.ColumnHeader;

            namespace MyNamespace;

            [InferColumnHeaders(typeof(ColumnHeaderResources), Prefix = "Header_")]
            public class ClassWithInferColumnHeaders
            {
                public string? FirstName { get; set; }
                public string? LastName { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithInferColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task InferColumnHeaders_ClassWithValidReferencesWithSuffix()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ColumnHeaders
            {
                public static string FirstNameHeader => "First name";
                public static string LastNameHeader => "Last name";
            }

            [InferColumnHeaders(typeof(ColumnHeaders), Suffix = "Header")]
            public class ClassWithInferColumnHeaders
            {
                public string? FirstName { get; set; }
                public string? LastName { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithInferColumnHeaders))]
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
                public string? {|SPCH1010:Name|} { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithInferColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }

    [Fact]
    public Task InferColumnHeaders_DerivedClassWithMissingReference()
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
            public abstract class BaseClassWithInferColumnHeaders
            {
            }

            public class DerivedClassWithInferColumnHeaders : BaseClassWithInferColumnHeaders
            {
                public string? {|SPCH1010:Name|} { get; set; }
            }
            
            [WorksheetRow(typeof(DerivedClassWithInferColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }

    [Fact]
    public Task InferColumnHeaders_ClassWithMissingReferenceForPropertyWithColumnIgnore()
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
                [ColumnIgnore]
                public string? Name { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithInferColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }

    [Fact]
    public Task InferColumnHeaders_DerivedClassWithMissingReferenceForPropertyWithColumnIgnore()
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
            public abstract class BaseClassWithInferColumnHeaders
            {
            }

            public class DerivedClassWithInferColumnHeaders : BaseClassWithInferColumnHeaders
            {
                [ColumnIgnore]
                public string? Name { get; set; }
            }
            
            [WorksheetRow(typeof(DerivedClassWithInferColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }

    [Fact]
    public Task InferColumnHeaders_ClassWithMissingReferenceForPropertyWithExplicitColumnHeader()
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
                [ColumnHeader("Explicit column header")]
                public string? Name { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithInferColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }

    [Fact]
    public Task InferColumnHeaders_ClassWithReferenceToInternalProperty()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ColumnHeaders
            {
                internal static string Name => "The name";
            }

            [InferColumnHeaders(typeof(ColumnHeaders))]
            public class ClassWithInferColumnHeaders
            {
                public string? {|SPCH1011:Name|} { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithInferColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }

    [Fact]
    public Task InferColumnHeaders_ClassWithReferenceToIntegerProperty()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ColumnHeaders
            {
                internal static int Name => 42;
            }

            [InferColumnHeaders(typeof(ColumnHeaders))]
            public class ClassWithInferColumnHeaders
            {
                public string? {|SPCH1004:Name|} { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithInferColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }
}
