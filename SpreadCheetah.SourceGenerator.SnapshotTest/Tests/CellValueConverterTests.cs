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
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;
            public class ClassWithInvalidConverter
            {
                [CellValueConverter(typeof(DecimalValueConverter))]
                public string? Property { get; set; }
            }
            
            internal class DecimalValueConverter : CellValueConverter<decimal>
            {
                public override DataCell ConvertToCell(decimal value) => new(value);
            }
            
            [WorksheetRow(typeof(ClassWithInvalidConverter))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source, onlyDiagnostics: true);
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
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;
            public class ClassWithConverterThatDoesNotInheritCellValueConverter
            {           
                [CellValueConverter(typeof(DecimalValueConverter))]
                public decimal Property { get; set; }
            }
            
            internal class DecimalValueConverter
            {
                public override DataCell ConvertToCell(decimal value) => new(value);
            }

            [WorksheetRow(typeof(ClassWithConverterThatDoesNotInheritCellValueConverter))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source, onlyDiagnostics: true);
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
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;
            public class ClassPropertyWithConverterAndCellValueTruncate
            {
                [CellValueConverter(typeof(StringValueConverter))]
                [CellValueTruncate(20)]
                public string? Property { get; set; }
            }

            internal class StringValueConverter : CellValueConverter<string>
            {
                public override DataCell ConvertToCell(string value) => new(value);
            }
               
            [WorksheetRow(typeof(ClassPropertyWithConverterAndCellValueTruncate))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source, onlyDiagnostics: true);
    }

    [Fact]
    public Task CellValueConverter_ClassWithoutParameterlessConstructor()
    {
        // Arrange
        const string source = """
            using SpreadCheetah.SourceGeneration;

            namespace MyNamespace;
            public class ClassWithoutParameterlessConstructor
            {
                [CellValueConverter(typeof(FixedValueConverter))]
                public string? Value { get; set; }
            }

            internal class FixedValueConverter(string fixedValue) : CellValueConverter<string>
            {
                public override DataCell ConvertToCell(string value) => new(fixedValue);
            }
               
            [WorksheetRow(typeof(ClassWithoutParameterlessConstructor))]
            public partial class MyGenRowContext : WorksheetRowContext;
            """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source, onlyDiagnostics: true);
    }
}