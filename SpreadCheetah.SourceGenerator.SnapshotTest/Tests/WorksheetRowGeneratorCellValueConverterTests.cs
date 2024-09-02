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
                              
                              [WorksheetRow(typeof(ClassWithCellValueConverters))]
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
                              using SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueConverters;
                              using System;
                              
                              namespace MyNamespace;
                              
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
                              using SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueConverters;
                              using System;

                              namespace MyNamespace;

                              [WorksheetRow(typeof(ClassWithInvalidCellValueConverter))]
                              public partial class MyGenRowContext : WorksheetRowContext;
                              """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
    
    [Fact]
    public Task WorksheetRowGenerator_When_Many_Rows_With_CellValueConverter_Should_Generate_One_CellValueDict()
    {

        // Arrange
        const string source = """
                              using SpreadCheetah.SourceGeneration;
                              using SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueConverters;
                              using System;

                              namespace MyNamespace;

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
                              using SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueConverters;
                              using System;

                              namespace MyNamespace;

                              [WorksheetRow(typeof(ClassWherePropertyTypeDifferentFromCellValueConverter))]
                              public partial class MyGenRowContext : WorksheetRowContext;
                              """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
    
    [Fact]
    public Task WorksheetRowGenerator_When_CellValueConverter_With_CellStyle_Should_Generate_Correctly()
    {

        // Arrange
        const string source = """
                              using SpreadCheetah.SourceGeneration;
                              using SpreadCheetah.SourceGenerator.SnapshotTest.Models.CellValueConverters;
                              using System;

                              namespace MyNamespace;

                              [WorksheetRow(typeof(ClassWithCellValueConvertersAndCellStyle))]
                              public partial class MyGenRowContext : WorksheetRowContext;
                              """;

        // Act & Assert
        return TestHelper.CompileAndVerify<WorksheetRowGenerator>(source);
    }
}