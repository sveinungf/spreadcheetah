using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGeneration.Internal;
using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class WorksheetRowGeneratorCellValueMapperTests
{
    [Fact]
    public Task WorksheetRowGenerator_Generate_RecordClassWith_2LevelOfInheritance_InheritedColumnsFirst_ParentIgnore_Inheritance()
    {

        // Arrange
        const string source = """
                              using SpreadCheetah.SourceGeneration;
                              using SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueMapping;
                              using System;
                              
                              namespace MyNamespace;
                              
                              [WorksheetRow(typeof(ClassWithCellValueMapper))]
                              public partial class MyGenRowContext : WorksheetRowContext;
                              """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
}