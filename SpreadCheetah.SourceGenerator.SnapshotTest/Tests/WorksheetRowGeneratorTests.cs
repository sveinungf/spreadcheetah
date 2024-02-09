using FluentAssertions;
using FluentAssertions.Execution;
using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class WorksheetRowGeneratorTests
{
    private static readonly string[] AllTrackingNames = ["InitialExtraction", "Transform"];

    [Fact]
    public void WorksheetRowGenerator_Generate_CachingCorrectly()
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

        const string expected = """... not shown for brevity...""";

        // Act
        var (diagnostics, output) = TestHelper.GetGeneratedTrees<WorksheetRowGenerator>([source], AllTrackingNames);

        // Assert the output
        using var s = new AssertionScope();
        diagnostics.Should().BeEmpty();
        output.Should().OnlyContain(x => x == expected);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithSingleProperty()
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
    public Task WorksheetRowGenerator_Generate_InternalClassWithSingleProperty()
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

    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithAllSupportedTypes()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(ClassWithAllSupportedTypes))]
                public partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_ClassWithMultipleProperties()
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
    public Task WorksheetRowGenerator_Generate_ClassWithNoProperties()
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
    public Task WorksheetRowGenerator_Generate_ClassWithNoPropertiesAndWarningsSuppressed()
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
    public Task WorksheetRowGenerator_Generate_ClassWithUnsupportedProperty()
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
    public Task WorksheetRowGenerator_Generate_ClassWithUnsupportedPropertyAndWarningsSuppressed()
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
    public Task WorksheetRowGenerator_Generate_ClassWithInheritance()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(ClassWithInheritance))]
                public partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_RecordClassWithSingleProperty()
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
    public Task WorksheetRowGenerator_Generate_StructWithSingleProperty()
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
    public Task WorksheetRowGenerator_Generate_RecordStructWithSingleProperty()
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
    public Task WorksheetRowGenerator_Generate_ReadOnlyStructWithSingleProperty()
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
    public Task WorksheetRowGenerator_Generate_ReadOnlyRecordStructWithSingleProperty()
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
    public Task WorksheetRowGenerator_Generate_RecordClassWithInheritance()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
            using System;

            namespace MyNamespace
            {
                [WorksheetRow(typeof(RecordClassWithInheritance))]
                public partial class MyGenRowContext : WorksheetRowContext
                {
                }
            }
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_ContextWithTwoWorksheetRowAttributes()
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
    public Task WorksheetRowGenerator_Generate_ContextWithTwoSimilarWorksheetRowAttributes()
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
    public Task WorksheetRowGenerator_Generate_ContextClassWithInternalAccessibility()
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
    public Task WorksheetRowGenerator_Generate_ContextClassWithDefaultAccessibility()
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
    public Task WorksheetRowGenerator_Generate_TwoContextClasses()
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
