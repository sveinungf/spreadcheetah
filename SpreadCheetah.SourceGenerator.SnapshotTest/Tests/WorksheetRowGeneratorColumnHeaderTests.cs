using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class WorksheetRowGeneratorColumnHeaderTests
{
    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithColumnHeaderForAllProperties()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models.ColumnHeader;

            namespace MyNamespace;
            
            [WorksheetRow(typeof(ClassWithColumnHeaderForAllProperties))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithSpecialCharacterColumnHeaders()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models.ColumnHeader;

            namespace MyNamespace;
            
            [WorksheetRow(typeof(ClassWithSpecialCharacterColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source, replaceEscapedLineEndings: true);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithPropertyReferenceColumnHeaders()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models.ColumnHeader;

            namespace MyNamespace;
            
            [WorksheetRow(typeof(ClassWithPropertyReferenceColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithInvalidPropertyReferenceColumnHeaders()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;

            public class ColumnHeaders
            {
                public static string Name => "The name";
                public static string PrivateGetterProperty { private get; set; } = "Private getter property";
                public static string WriteOnlyProperty { set => _ = value; }
                public static int NonStringProperty => 2024;
                internal static string InternalProperty => "Internal property";
                public string NonStaticProperty => "Non static property";
            }

            public class ClassWithInvalidPropertyReferenceColumnHeaders
            {
                [ColumnHeader(typeof(ColumnHeaders), "NonExistingProperty")]
                public string PropertyA { get; set; }
                [ColumnHeader(typeof(ColumnHeaders), "name")]
                public string PropertyB { get; set; }
                [ColumnHeader(typeof(ColumnHeaders), nameof(ColumnHeaders.PrivateGetterProperty))]
                public string PropertyC { get; set; }
                [ColumnHeader(typeof(ColumnHeaders), nameof(ColumnHeaders.WriteOnlyProperty))]
                public string PropertyD { get; set; }
                [ColumnHeader(typeof(ColumnHeaders), nameof(ColumnHeaders.NonStringProperty))]
                public string PropertyE { get; set; }
                [ColumnHeader(typeof(ColumnHeaders), nameof(ColumnHeaders.InternalProperty))]
                public string PropertyF { get; set; }
                [ColumnHeader(typeof(ColumnHeaders), nameof(ColumnHeaders.NonStaticProperty))]
                public string PropertyG { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithInvalidPropertyReferenceColumnHeaders))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
}
