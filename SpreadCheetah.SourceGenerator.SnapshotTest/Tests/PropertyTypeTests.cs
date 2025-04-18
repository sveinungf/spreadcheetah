using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class PropertyTypeTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public Task PropertyType_ClassWithUnsupportedProperty()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;
            using System;

            namespace MyNamespace;

            public class ClassWithUnsupportedProperty
            {
                public System.Diagnostics.Stopwatch? Stopwatch { get; set; }
            }
            
            [WorksheetRow({|SPCH1002:typeof(ClassWithUnsupportedProperty)|})]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }

    [Fact]
    public Task PropertyType_ClassWithUnsupportedPropertyAndWarningsSuppressed()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah.SourceGeneration;
            using System;

            namespace MyNamespace;
            
            public class ClassWithUnsupportedProperty
            {
                public System.Diagnostics.Stopwatch? Stopwatch { get; set; }
            }
            
            [WorksheetRow(typeof(ClassWithUnsupportedProperty))]
            [WorksheetRowGenerationOptions(SuppressWarnings = true)]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync(Token);
    }

    [Fact]
    public Task PropertyType_ClassWithAllSupportedTypes()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using System;

            namespace MyNamespace;

            public class ClassWithAllSupportedTypes
            {
                public string StringValue { get; set; } = "";
                public string? NullableStringValue { get; set; }
                public int IntValue { get; set; }
                public int? NullableIntValue { get; set; }
                public long LongValue { get; set; }
                public long? NullableLongValue { get; set; }
                public float FloatValue { get; set; }
                public float? NullableFloatValue { get; set; }
                public double DoubleValue { get; set; }
                public double? NullableDoubleValue { get; set; }
                public decimal DecimalValue { get; set; }
                public decimal? NullableDecimalValue { get; set; }
                public DateTime DateTimeValue { get; set; }
                public DateTime? NullableDateTimeValue { get; set; }
                public bool BoolValue { get; set; }
                public bool? NullableBoolValue { get; set; }
            }

            [WorksheetRow(typeof(ClassWithAllSupportedTypes))]
            public partial class MyContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task PropertyType_ClassWithIndexer()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            using System;

            namespace MyNamespace;

            public class MyClass
            {
                public int this[int index] => index;
            }

            [WorksheetRow(typeof(MyClass))]
            public partial class MyContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
}
