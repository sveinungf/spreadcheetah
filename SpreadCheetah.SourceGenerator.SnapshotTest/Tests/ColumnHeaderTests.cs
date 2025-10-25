using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class ColumnHeaderTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public Task ColumnHeader_ClassWithColumnHeaderForAllProperties()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithColumnHeaderForAllProperties
            {
                [ColumnHeader("First name")]
                public string FirstName { get; set; } = "";

                [ColumnHeader("Middle name")]
                public string? MiddleName { get; set; }

                [ColumnHeader("Last name")]
                public string LastName { get; set; } = "";

                [ColumnHeader("Age")]
                public int Age { get; set; }

                [ColumnHeader("Employed (yes/no)")]
                public bool Employed { get; set; }

                [ColumnHeader("Score (decimal)")]
                public double Score { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithColumnHeaderForAllProperties))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task ColumnHeader_ClassWithSpecialCharacterColumnHeaders()
    {
        // Arrange
        const string source = """"
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ClassWithSpecialCharacterColumnHeaders
            {
                [ColumnHeader("First name")]
                public string? FirstName { get; set; }

                [ColumnHeader("")]
                public string? LastName { get; set; }

                [ColumnHeader("Nationality (escaped characters \", \', \\)")]
                public string? Nationality { get; set; }

                [ColumnHeader("Address line 1 (escaped characters \r\n, \t)")]
                public string? AddressLine1 { get; set; }

                [ColumnHeader(@"Address line 2 (verbatim
            string: "", \)")]
                public string? AddressLine2 { get; set; }

                [ColumnHeader("""
                    Age (
                        raw
                        string
                        literal
                    )
                """)]
                public int Age { get; set; }

                [ColumnHeader("Note (unicode escape sequence 🌉, \ud83d\udc4d, \xE7)")]
                public string? Note { get; set; }

                private const string Constant = "This is a constant";

                [ColumnHeader($"Note 2 (constant interpolated string: {Constant})")]
                public string? Note2 { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithSpecialCharacterColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """";

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source, replaceEscapedLineEndings: true);
    }

    [Fact]
    public Task ColumnHeader_ClassWithValidPropertiesForColumnHeaderReference()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models.ColumnHeader;

            namespace MyNamespace;

            public class ClassWithPropertyReferenceColumnHeaders
            {
                [ColumnHeader(typeof(ColumnHeaderResources), nameof(ColumnHeaderResources.Header_FirstName))]
                public string? FirstName { get; set; }

                [ColumnHeader(propertyName: nameof(ColumnHeaderResources.Header_LastName), type: typeof(ColumnHeaderResources))]
                public string? LastName { get; set; }

                [ColumnHeader(typeof(ColumnHeaders), nameof(ColumnHeaders.HeaderNationality))]
                public string? Nationality { get; set; }

                [ColumnHeader(typeof(ColumnHeaders), nameof(ColumnHeaders.HeaderAddressLine1))]
                public string? AddressLine1 { get; set; }

                [ColumnHeader(typeof(ColumnHeaders), nameof(ColumnHeaders.HeaderAddressLine2))]
                public string? AddressLine2 { get; set; }

                [ColumnHeader(typeof(ColumnHeaders), nameof(ColumnHeaders.HeaderAge))]
                public int Age { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithPropertyReferenceColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task ColumnHeader_ClassWithMissingPropertiesForColumnHeaderReference()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ColumnHeaders
            {
                public static string Name => "The name";
            }

            public class ClassWithInvalidPropertyReferenceColumnHeaders
            {
                [ColumnHeader{|SPCH1010:(typeof(ColumnHeaders), "NonExistingProperty")|}]
                public string PropertyA { get; set; }
                [ColumnHeader{|SPCH1010:(typeof(ColumnHeaders), "name")|}]
                public string PropertyB { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithInvalidPropertyReferenceColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }

    [Fact]
    public Task ColumnHeader_ClassWithUnsupportedPropertiesForColumnHeaderReference()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ColumnHeaders
            {
                public static string PrivateGetterProperty { private get; set; } = "Private getter property";
                public static string WriteOnlyProperty { set => _ = value; }
                public static int NonStringProperty => 2024;
                internal static string InternalProperty => "Internal property";
                public string NonStaticProperty => "Non static property";
            }

            public class ClassWithInvalidPropertyReferenceColumnHeaders
            {
                [ColumnHeader{|SPCH1004:(typeof(ColumnHeaders), nameof(ColumnHeaders.PrivateGetterProperty))|}]
                public string PropertyC { get; set; }
                [ColumnHeader{|SPCH1004:(typeof(ColumnHeaders), nameof(ColumnHeaders.WriteOnlyProperty))|}]
                public string PropertyD { get; set; }
                [ColumnHeader{|SPCH1004:(typeof(ColumnHeaders), nameof(ColumnHeaders.NonStringProperty))|}]
                public string PropertyE { get; set; }
                [ColumnHeader{|SPCH1004:(typeof(ColumnHeaders), nameof(ColumnHeaders.InternalProperty))|}]
                public string PropertyF { get; set; }
                [ColumnHeader{|SPCH1004:(typeof(ColumnHeaders), nameof(ColumnHeaders.NonStaticProperty))|}]
                public string PropertyG { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithInvalidPropertyReferenceColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }
}
