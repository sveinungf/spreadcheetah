using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class WorksheetRowGeneratorTests
{
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
}
