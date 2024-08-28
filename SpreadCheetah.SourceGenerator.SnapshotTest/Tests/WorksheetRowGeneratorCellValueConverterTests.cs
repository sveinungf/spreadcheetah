using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGeneration.Internal;
using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class WorksheetRowGeneratorCellValueConverterTests
{
    [Fact]
    public Task WorksheetRowGenerator_Generate_Class_With_Different_Converters()
    {

        // Arrange
        const string source = """
                              using SpreadCheetah.SourceGeneration;
                              using SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueConverters;
                              using System;
                              
                              namespace MyNamespace;
                              
                              [WorksheetRow(typeof(ClassWithCellValueConverters))]
                              public partial class MyGenRowContext : WorksheetRowContext;
                              """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
}