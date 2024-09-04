using SpreadCheetah.SourceGenerator.SnapshotTest.Helpers;
using SpreadCheetah.SourceGenerators;

namespace SpreadCheetah.SourceGenerator.SnapshotTest.Tests;

public class WorksheetRowGeneratorCellValueConverterTests
{
    [Fact]
    public Task WorksheetRowGenerator_Generate_Class_With_Different_CellValueConverters_Should_Generate_All_Converters()
    {

        // Arrange
        const string source = """
                              using SpreadCheetah.SourceGeneration;
                              using SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueConverters;
                              using System;
                              
                              namespace MyNamespace;
                              public class ClassWithSameCellValueConverters
                              {
                                  [ColumnHeader("Property")]
                                  [CellValueConverter(typeof(StringValueConverter))]
                                  public string? Property { get; set; }
                                  
                                  [ColumnHeader("Property1")]
                                  [CellValueConverter(typeof(NullableIntValueConverter))]
                                  public int? Property1 { get; set; }
                                  
                                  [ColumnHeader("Property1")]
                                  [CellValueConverter(typeof(DecimalValueConverter))]
                                  public decimal Property2 { get; set; }
                              }
                              
                              internal class StringValueConverter : CellValueConverter<string>
                              {
                                  public override DataCell ConvertToCell(string value)
                                  {
                                      return new DataCell(value);
                                  }
                              }
                              
                              internal class NullableIntValueConverter : CellValueConverter<int?>
                              {
                                  public override DataCell ConvertToCell(int? value)
                                  {
                                      return new DataCell(value);
                                  }
                              }
                              
                              internal class DecimalValueConverter : CellValueConverter<decimal>
                              {
                                  public override DataCell ConvertToCell(decimal value)
                                  {
                                      return new DataCell(value);
                                  }
                              }
                              
                              [WorksheetRow(typeof(ClassWithSameCellValueConverters))]
                              public partial class MyGenRowContext : WorksheetRowContext;
                              """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
    
    [Fact]
    public Task WorksheetRowGenerator_Generate_Class_With_Same_CellValueConverters_Should_Generate_Unique_Converters()
    {

        // Arrange
        const string source = """
                              using SpreadCheetah.SourceGeneration;
                              using System;
                              
                              namespace MyNamespace;
                              public class ClassWithSameCellValueConverters
                              {
                                  [ColumnHeader("Property")]
                                  [CellValueConverter(typeof(StringValueConverter))]
                                  public string? Property { get; set; }
                                  
                                  [ColumnHeader("Property1")]
                                  [CellValueConverter(typeof(StringValueConverter))]
                                  public string? Property1 { get; set; }
                                  
                                  [ColumnHeader("Property1")]
                                  [CellValueConverter(typeof(DecimalValueConverter))]
                                  public decimal Property2 { get; set; }
                              }
                              
                              internal class StringValueConverter : CellValueConverter<string>
                              {
                                  public override DataCell ConvertToCell(string value)
                                  {
                                      return new DataCell(value);
                                  }
                              }
                              
                              internal class DecimalValueConverter : CellValueConverter<decimal>
                              {
                                  public override DataCell ConvertToCell(decimal value)
                                  {
                                      return new DataCell(value);
                                  }
                              }
                              
                              [WorksheetRow(typeof(ClassWithSameCellValueConverters))]
                              public partial class MyGenRowContext : WorksheetRowContext;
                              """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
    
    [Fact]
    public Task WorksheetRowGenerator_When_Class_Have_Invalid_CellValueConverter_Should_Emit_Error()
    {

        // Arrange
        const string source = """
                              using SpreadCheetah.SourceGeneration;
                              using System;

                              namespace MyNamespace;
                              public class ClassWithInvalidCellValueConverter
                              {
                                  [ColumnHeader("Property")]
                                  [CellValueConverter(typeof(DecimalValueConverter))]
                                  public string? Property { get; set; }
                                  
                                  [ColumnHeader("Property1")]
                                  [CellValueConverter(typeof(DecimalValueConverter))]
                                  public string? Property1 { get; set; }
                              }
                              
                              internal class DecimalValueConverter : CellValueConverter<decimal>
                              {
                                  public override DataCell ConvertToCell(decimal value)
                                  {
                                      return new DataCell(value);
                                  }
                              }
                              
                              
                              [WorksheetRow(typeof(ClassWithInvalidCellValueConverter))]
                              public partial class MyGenRowContext : WorksheetRowContext;
                              """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source, onlyDiagnostics: true);
    }
    
    [Fact]
    public Task WorksheetRowGenerator_When_Many_Rows_With_CellValueConverter_Should_Generate_One_Converter_Only_Once()
    {

        // Arrange
        const string source = """
                              using SpreadCheetah.SourceGeneration;
                              using System;

                              namespace MyNamespace;
                              public class ClassWithSameCellValueConverters
                              {
                                  [ColumnHeader("Property")]
                                  [CellValueConverter(typeof(StringValueConverter))]
                                  public string? Property { get; set; }
                                  
                                  [ColumnHeader("Property1")]
                                  [CellValueConverter(typeof(StringValueConverter))]
                                  public string? Property1 { get; set; }
                                  
                                  [ColumnHeader("Property1")]
                                  [CellValueConverter(typeof(DecimalValueConverter))]
                                  public decimal Property2 { get; set; }
                              }
                              
                              public class ClassWithCellValueConvertersAndCellStyle
                              {
                                  [ColumnHeader("Property")]
                                  [CellValueConverter(typeof(StringValueConverter))]
                                  [CellStyle("Test")]
                                  public string? Property { get; set; }
                                  
                                  [ColumnHeader("Property1")]
                                  [CellValueConverter(typeof(NullableIntValueConverter))]
                                  public int? Property1 { get; set; }
                                  
                                  [ColumnHeader("Property1")]
                                  public decimal Property2 { get; set; }
                              }
                              
                              internal class StringValueConverter : CellValueConverter<string>
                              {
                                  public override DataCell ConvertToCell(string value)
                                  {
                                      return new DataCell(value);
                                  }
                              }
                              
                              internal class DecimalValueConverter : CellValueConverter<decimal>
                              {
                                  public override DataCell ConvertToCell(decimal value)
                                  {
                                      return new DataCell(value);
                                  }
                              }
                              
                              [WorksheetRow(typeof(ClassWithCellValueConverters))]
                              [WorksheetRow(typeof(ClassWithSameCellValueConverters))]
                              public partial class MyGenRowContext : WorksheetRowContext;
                              """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
    
    [Fact]
    public Task WorksheetRowGenerator_When_Property_Type_Not_Same_As_CellValueConverter()
    {

        // Arrange
        const string source = """
                              using SpreadCheetah.SourceGeneration;
                              using System;

                              namespace MyNamespace;
                              public class ClassWherePropertyTypeDifferentFromCellValueConverter
                              {
                                  [CellValueConverter(typeof(NullableIntValueConverter))]
                                  public string Property { get; set; } = null!;
                                  
                                  [CellValueConverter(typeof(DecimalValueConverter))]
                                  public int? Property1 { get; set; }
                              }
                              
                              internal class NullableIntValueConverter : CellValueConverter<int?>
                              {
                                  public override DataCell ConvertToCell(int? value)
                                  {
                                      return new DataCell(value);
                                  }
                              }
                              
                              internal class DecimalValueConverter : CellValueConverter<decimal>
                              {
                                  public override DataCell ConvertToCell(decimal value)
                                  {
                                      return new DataCell(value);
                                  }
                              }
                              
                              [WorksheetRow(typeof(ClassWherePropertyTypeDifferentFromCellValueConverter))]
                              public partial class MyGenRowContext : WorksheetRowContext;
                              """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source, onlyDiagnostics: true);
    }
    
    [Fact]
    public Task WorksheetRowGenerator_When_Class_Doesnt_Inherit_CellValueConverter_Should_EmitError()
    {

        // Arrange
        const string source = """
                              using SpreadCheetah.SourceGeneration;
                              using System;

                              namespace MyNamespace;
                              public class ClassWherePropertyTypeDifferentFromCellValueConverter
                              {
                                  [CellValueConverter(typeof(NullableIntValueConverter))]
                                  public string Property { get; set; } = null!;
                                  
                                  [CellValueConverter(typeof(DecimalValueConverter))]
                                  public int? Property1 { get; set; }
                              }
                              
                              internal class NullableIntValueConverter : CellValueConverter<int?>
                              {
                                  public override DataCell ConvertToCell(int? value)
                                  {
                                      return new DataCell(value);
                                  }
                              }
                              
                              internal class DecimalValueConverter : CellValueConverter<decimal>
                              {
                                  public override DataCell ConvertToCell(decimal value)
                                  {
                                      return new DataCell(value);
                                  }
                              }

                              [WorksheetRow(typeof(ClassWherePropertyTypeDifferentFromCellValueConverter))]
                              public partial class MyGenRowContext : WorksheetRowContext;
                              """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source, onlyDiagnostics: true);
    }
    
    [Fact]
    public Task WorksheetRowGenerator_When_CellValueConverter_With_CellStyle_Should_Generate_Correctly()
    {

        // Arrange
        const string source = """
                              using SpreadCheetah.SourceGeneration;
                              using System;

                              namespace MyNamespace;
                              public class ClassWithCellValueConvertersAndCellStyle
                              {
                                  [ColumnHeader("Property")]
                                  [CellValueConverter(typeof(StringValueConverter))]
                                  [CellStyle("Test")]
                                  public string? Property { get; set; }
                                  
                                  [ColumnHeader("Property1")]
                                  [CellValueConverter(typeof(NullableIntValueConverter))]
                                  public int? Property1 { get; set; }
                                  
                                  [ColumnHeader("Property1")]
                                  public decimal Property2 { get; set; }
                              }

                              internal class StringValueConverter : CellValueConverter<string>
                              {
                                  public override DataCell ConvertToCell(string value)
                                  {
                                      return new DataCell(value);
                                  }
                              }
                              
                              internal class NullableIntValueConverter : CellValueConverter<int?>
                              {
                                  public override DataCell ConvertToCell(int? value)
                                  {
                                      return new DataCell(value);
                                  }
                              }
                              
                              [WorksheetRow(typeof(ClassWithCellValueConvertersAndCellStyle))]
                              public partial class MyGenRowContext : WorksheetRowContext;
                              """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
}