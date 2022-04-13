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
    public Task WorksheetRowSourceGenerator_Generate_InternalClassWithSingleProperty()
    {
        // Arrange
        var source = @"
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
    public Task WorksheetRowSourceGenerator_Generate_ClassWithUnsupportedProperty()
    {
        // Arrange
        var source = @"
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
using System;

namespace MyNamespace
{
    [WorksheetRow(typeof(ClassWithUnsupportedProperty))]
    public partial class MyGenRowContext : WorksheetRowGeneratorContext
    {
    }
}";

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowSourceGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ClassWithUnsupportedPropertyAndWarningsSuppressed()
    {
        // Arrange
        var source = @"
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
using System;

namespace MyNamespace
{
    [WorksheetRow(typeof(ClassWithUnsupportedProperty))]
    [WorksheetRowGenerationOptions(SuppressWarnings = true)]
    public partial class MyGenRowContext : WorksheetRowGeneratorContext
    {
    }
}";

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowSourceGenerator>(source);
    }

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_ContextWithTwoWorksheetRowAttributes()
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
    public Task WorksheetRowSourceGenerator_Generate_ContextWithTwoSimilarWorksheetRowAttributes()
    {
        // Arrange
        var source = @"
using SpreadCheetah.SourceGeneration;
using System;

namespace MyNamespace
{
    [WorksheetRow(typeof(SpreadCheetah.SourceGenerator.SnapshotTest.Models.ClassWithSingleProperty))]
    [WorksheetRow(typeof(SpreadCheetah.SourceGenerator.SnapshotTest.AlternativeModels.ClassWithSingleProperty))]
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

    [Fact]
    public Task WorksheetRowSourceGenerator_Generate_TwoContextClasses()
    {
        // Arrange
        var source = @"
using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.SnapshotTest.Models;
using System;

namespace MyNamespace;

[WorksheetRow(typeof(ClassWithSingleProperty))]
public partial class MyGenRowContext : WorksheetRowGeneratorContext
{
}

[WorksheetRow(typeof(ClassWithMultipleProperties))]
public partial class MyGenRowContext2 : WorksheetRowGeneratorContext
{
}
";

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowSourceGenerator>(source);
    }
}
