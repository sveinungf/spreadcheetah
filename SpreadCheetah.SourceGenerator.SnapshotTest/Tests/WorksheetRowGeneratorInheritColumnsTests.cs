using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class WorksheetRowGeneratorInheritColumnsTests
{
    [Fact]
    public Task WorksheetRowGenerator_Generate_RecordClassWith_2LevelOfInheritance_InheritedColumnsFirst_ParentIgnore_Inheritance()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models.InheritColumns;
            using System;

            namespace MyNamespace;
            
            [WorksheetRow(typeof(RecordClass2LevelOfInheritanceInheritedColumnsFirstParentIgnoreInheritance))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_RecordClassWith_2LevelOfInheritance_InheritedColumnsFirst()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models.InheritColumns;
            using System;

            namespace MyNamespace;

            [WorksheetRow(typeof(RecordClass2LevelOfInheritanceInheritedColumnsFirst))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_RecordClassWith_2LevelOfInheritance_InheritedColumnsLast_ParentIgnore_Inheritance()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models.InheritColumns;
            using System;

            namespace MyNamespace;
                              
            [WorksheetRow(typeof(RecordClass2LevelOfInheritanceInheritedColumnsLastParentIgnoreInheritance))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_RecordClassWith_2LevelOfInheritance_InheritedColumnsLast()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models.InheritColumns;
            using System;

            namespace MyNamespace;
                              
            [WorksheetRow(typeof(RecordClass2LevelOfInheritanceInheritedColumnsLast))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_RecordClassWith_2LevelOfInheritance_InheritedColumnsLast_ParentInheritColumnsLast()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models.InheritColumns;
            using System;

            namespace MyNamespace;
                              
            [WorksheetRow(typeof(RecordClass2LevelOfInheritanceInheritedColumnsFirstParentInheritColumnsLast))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowGenerator_Generate_RecordClassWith_2LevelOfInheritance_InheritedColumnsLast_ParentInheritColumnsFirst()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using SpreadCheetah.SourceGenerator.SnapshotTest.Models.InheritColumns;
            using System;

            namespace MyNamespace;
                              
            [WorksheetRow(typeof(RecordClass2LevelOfInheritanceInheritedColumnsLastParentInheritColumnsFirst))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
}
