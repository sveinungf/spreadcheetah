using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class CellValueConverterTests
{
    [Fact]
    public Task CellValueConverter_ClassWithMultipleConverters()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            
            namespace MyNamespace;
            public class ClassWithMultipleConverters
            {
                [CellValueConverter(typeof(StringValueConverter))]
                public string? Property { get; set; }

                [CellValueConverter(typeof(NullableIntValueConverter))]
                public int? Property1 { get; set; }

                [CellValueConverter(typeof(DecimalValueConverter))]
                public decimal Property2 { get; set; }
            }
            
            internal class StringValueConverter : CellValueConverter<string>
            {
                public override DataCell ConvertToCell(string value) => new(value);
            }
            
            internal class NullableIntValueConverter : CellValueConverter<int?>
            {
                public override DataCell ConvertToCell(int? value) => new(value);
            }
            
            internal class DecimalValueConverter : CellValueConverter<decimal>
            {
                public override DataCell ConvertToCell(decimal value) => new(value);
            }
            
            [WorksheetRow(typeof(ClassWithMultipleConverters))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task CellValueConverter_ClassWithReusedConverter()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;
            
            namespace MyNamespace;
            public class ClassWithReusedConverter
            {
                [CellValueConverter(typeof(StringValueConverter))]
                public string? Property1 { get; set; }

                [CellValueConverter(typeof(DecimalValueConverter))]
                public decimal Property2 { get; set; }

                [CellValueConverter(typeof(StringValueConverter))]
                public string? Property3 { get; set; }
            }
            
            internal class StringValueConverter : CellValueConverter<string>
            {
                public override DataCell ConvertToCell(string value) => new(value);
            }

            internal class DecimalValueConverter : CellValueConverter<decimal>
            {
                public override DataCell ConvertToCell(decimal value) => new(value);
            }
            
            [WorksheetRow(typeof(ClassWithReusedConverter))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task CellValueConverter_ClassWithInvalidConverter()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah;
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;
            public class ClassWithInvalidConverter
            {
                [CellValueConverter({|SPCH1007:typeof(DecimalValueConverter)|})]
                public string? Property { get; set; }
            }
            
            internal class DecimalValueConverter : CellValueConverter<decimal>
            {
                public override DataCell ConvertToDataCell(decimal value) => new(value);
            }
            
            [WorksheetRow(typeof(ClassWithInvalidConverter))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync();
    }

    [Fact]
    public Task CellValueConverter_TwoClassesUsingTheSameConverter()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;
            public class Person
            {
                [CellValueConverter(typeof(StringValueConverter))]
                public string? Name { get; set; }
            }
                              
            public class Car
            {
                [CellValueConverter(typeof(StringValueConverter))]
                public string? Model { get; set; }
            }
                              
            internal class StringValueConverter : CellValueConverter<string>
            {
                public override DataCell ConvertToCell(string value) => new(value);
            }
             
            [WorksheetRow(typeof(Person))]
            [WorksheetRow(typeof(Car))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task CellValueConverter_ClassWithConverterThatDoesNotInheritCellValueConverter()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah;
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;
            public class ClassWithConverterThatDoesNotInheritCellValueConverter
            {           
                [CellValueConverter({|SPCH1007:typeof(DecimalValueConverter)|})]
                public decimal Property { get; set; }
            }
            
            internal class DecimalValueConverter
            {
                public DataCell ConvertToDataCell(decimal value) => new(value);
            }

            [WorksheetRow(typeof(ClassWithConverterThatDoesNotInheritCellValueConverter))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync();
    }

    [Fact]
    public Task CellValueConverter_ClassPropertyWithConverterAndCellStyle()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;
            public class ClassPropertyWithConverterAndCellStyle
            {
                [CellValueConverter(typeof(StringValueConverter))]
                [CellStyle("My style")]
                public string? Property { get; set; }
            }

            internal class StringValueConverter : CellValueConverter<string>
            {
                public override DataCell ConvertToCell(string value) => new(value);
            }
               
            [WorksheetRow(typeof(ClassPropertyWithConverterAndCellStyle))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task CellValueConverter_ClassPropertyWithConverterAndCellValueTruncate()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah;
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;
            public class ClassPropertyWithConverterAndCellValueTruncate
            {
                [CellValueConverter(typeof(StringValueConverter))]
                [{|SPCH1008:CellValueTruncate(20)|}]
                public string? Property { get; set; }
            }

            internal class StringValueConverter : CellValueConverter<string>
            {
                public override DataCell ConvertToDataCell(string value) => new(value);
            }
               
            [WorksheetRow(typeof(ClassPropertyWithConverterAndCellValueTruncate))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync();
    }

    [Fact]
    public Task CellValueConverter_ClassWithoutParameterlessConstructor()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah;
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;
            public class ClassWithoutParameterlessConstructor
            {
                [CellValueConverter({|SPCH1009:typeof(FixedValueConverter)|})]
                public string? Value { get; set; }
            }

            internal class FixedValueConverter(string fixedValue) : CellValueConverter<string>
            {
                public override DataCell ConvertToDataCell(string value) => new(fixedValue);
            }
               
            [WorksheetRow(typeof(ClassWithoutParameterlessConstructor))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync();
    }

    [Fact]
    public Task CellValueConverter_ClassWithConverterOnComplexProperty()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;
            public class ClassWithConverterOnComplexProperty
            {
                [CellValueConverter(typeof(ObjectConverter))]
                public object? Property1 { get; set; }
                
                [CellValueConverter(typeof(ObjectConverter))]
                [CellStyle("object style")]
                public object? Property2 { get; set; }
            }

            internal class ObjectConverter : CellValueConverter<object>
            {
                public override DataCell ConvertToCell(object value) => new(value.ToString());
            } 
                  
            [WorksheetRow(typeof(ClassWithConverterOnComplexProperty))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }

    [Fact]
    public Task CellValueConverter_ClassWithInvalidConverterOnComplexProperty()
    {
        // Arrange
        var context = AnalyzerTest.CreateContext();
        context.TestCode = """
            using SpreadCheetah;
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;
            public class ClassWithInvalidConverterOnComplexProperty
            {
                [CellValueConverter({|SPCH1007:typeof(StringConverter)|})]
                public object? Property1 { get; set; }

                public string? Property2 { get; set; }
            }
               
            internal class StringConverter : CellValueConverter<string>
            {
                public override DataCell ConvertToDataCell(string value) => new(value);
            } 
                                 
            [WorksheetRow(typeof(ClassWithInvalidConverterOnComplexProperty))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return context.RunAsync();
    }
}