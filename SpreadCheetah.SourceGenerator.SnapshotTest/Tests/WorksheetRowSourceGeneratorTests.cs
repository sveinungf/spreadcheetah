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

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ClassWithMultipleProperties()
    {
        // Arrange
        var source = @"
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
using System;

namespace MyNamespace
{
    [WorksheetRow(typeof(ClassWithMultipleProperties))]
    public partial class MyGenRowContext : WorksheetRowGeneratorContext
    {
    }
}";

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowSourceGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ClassWithNoProperties()
    {
        // Arrange
        var source = @"
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
using System;

namespace MyNamespace
{
    [WorksheetRow(typeof(ClassWithNoProperties))]
    public partial class MyGenRowContext : WorksheetRowGeneratorContext
    {
    }
}";

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowSourceGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ClassWithNoPropertiesAndWarningsSuppressed()
    {
        // Arrange
        var source = @"
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
using System;

namespace MyNamespace
{
    [WorksheetRowGenerationOptions(SuppressWarnings = true)]
    [WorksheetRow(typeof(ClassWithNoProperties))]
    public partial class MyGenRowContext : WorksheetRowGeneratorContext
    {
    }
}";

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowSourceGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ContextWithMultipleWorksheetRowAttributes()
    {
        // Arrange
        var source = @"
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
using System;

namespace MyNamespace
{
    [WorksheetRow(typeof(ClassWithSingleProperty))]
    [WorksheetRow(typeof(ClassWithMultipleProperties))]
    public partial class MyGenRowContext : WorksheetRowGeneratorContext
    {
    }
}";

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowSourceGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ContextClassWithInternalAccessibility()
    {
        // Arrange
        var source = @"
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
using System;

namespace MyNamespace
{
    [WorksheetRow(typeof(ClassWithSingleProperty))]
    internal partial class MyGenRowContext : WorksheetRowGeneratorContext
    {
    }
}";

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowSourceGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ContextClassWithDefaultAccessibility()
    {
        // Arrange
        var source = @"
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
using System;

namespace MyNamespace
{
    [WorksheetRow(typeof(ClassWithSingleProperty))]
    partial class MyGenRowContext : WorksheetRowGeneratorContext
    {
    }
}";

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowSourceGenerator>(source);
    }
}
