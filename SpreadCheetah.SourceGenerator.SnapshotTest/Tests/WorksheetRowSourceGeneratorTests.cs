using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

[UsesVerify]
public class WorksheetRowSourceGeneratorTests
{
    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ClassWithSingleProperty()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(ClassWithSingleProperty))]
                public partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_InternalClassWithSingleProperty()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using System;

            namespace MyNamespace
            {
                internal class InternalClassWithSingleProperty
                {
                    public string? Name { get; set; }
                }

                [WorksheetRow(typeof(InternalClassWithSingleProperty))]
                internal partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    // TODO: Fill in all supported types
    [Theory]
    [InlineData(CellValueType.Bool, false)]
    [InlineData(CellValueType.Int, false)]
    [InlineData(CellValueType.Int, true)]
    public Task WorksheetRowSourceGenerator_Generate_ClassWithSingleGenericProperty(CellValueType type, bool nullable)
    {
        var keyword = type switch
        {
            CellValueType.Bool => "bool",
            CellValueType.Int => "int",
            _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
        };

        if (nullable)
            keyword += "?";

        // Arrange
        var source = $$"""
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(ClassWithSingleGenericProperty<{{keyword}}>))]
                public partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source, type, nullable);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ClassWithMultipleProperties()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(ClassWithMultipleProperties))]
                public partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ClassWithNoProperties()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(ClassWithNoProperties))]
                public partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ClassWithNoPropertiesAndWarningsSuppressed()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRowGenerationOptions(SuppressWarnings = true)]
                [WorksheetRow(typeof(ClassWithNoProperties))]
                public partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ClassWithUnsupportedProperty()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(ClassWithUnsupportedProperty))]
                public partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ClassWithUnsupportedPropertyAndWarningsSuppressed()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(ClassWithUnsupportedProperty))]
                [WorksheetRowGenerationOptions(SuppressWarnings = true)]
                public partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_RecordClassWithSingleProperty()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(RecordClassWithSingleProperty))]
                public partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_StructWithSingleProperty()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(StructWithSingleProperty))]
                public partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_RecordStructWithSingleProperty()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(RecordStructWithSingleProperty))]
                public partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ReadOnlyStructWithSingleProperty()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(ReadOnlyStructWithSingleProperty))]
                public partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ReadOnlyRecordStructWithSingleProperty()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(ReadOnlyRecordStructWithSingleProperty))]
                public partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ContextWithTwoWorksheetRowAttributes()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(ClassWithSingleProperty))]
                [WorksheetRow(typeof(ClassWithMultipleProperties))]
                public partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ContextWithTwoSimilarWorksheetRowAttributes()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithSingleProperty))]
                [WorksheetRow(typeof(SpreadCheetah.SourceGenerator.SnapshotTest.AlternativeModels.ClassWithSingleProperty))]
                public partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ContextClassWithInternalAccessibility()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(ClassWithSingleProperty))]
                internal partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ContextClassWithDefaultAccessibility()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(ClassWithSingleProperty))]
                partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_TwoContextClasses()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace;

            [WorksheetRow(typeof(ClassWithSingleProperty))]
            public partial class MyGenRowContext : WorksheetRowContext
            {
            }

            [WorksheetRow(typeof(ClassWithMultipleProperties))]
            public partial class MyGenRowContext2 : WorksheetRowContext
            {
            }

            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
}
