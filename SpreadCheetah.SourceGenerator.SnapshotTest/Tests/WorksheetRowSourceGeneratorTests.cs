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
        var source = @"
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
using System;

namespace MyNamespace
{
    [WorksheetRow(typeof(ClassWithSingleProperty))]
    public partial class MyGenRowContext : WorksheetRowGeneratorContext
    {
    }
}";

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowSourceGenerator>(source);
    }
}
