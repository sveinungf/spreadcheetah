using SpreadCheetah.SourceGenerator.Test.Models.CellValueConverter;
using SpreadCheetah.TestHelpers.Assertions;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class CellAndColumnAttributeTests
{
    //
    // CellFormatAttribute
    //

    [Fact]
    public async Task CellFormatAttribute_ClassWithReusedConverter()
    {
        //
        //// Arrange
        //using var stream = new MemoryStream();
        //await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        //await spreadsheet.StartWorksheetAsync("Sheet");
        //var obj = new ClassWithReusedConverter
        //{
        //    FirstName = "Ola",
        //    MiddleName = null,
        //    LastName = "Nordmann",
        //    Gpa = 3.1m
        //};

        //// Act
        //await spreadsheet.AddAsRowAsync(obj, CellValueConverterContext.Default.ClassWithReusedConverter);
        //await spreadsheet.FinishAsync();

        //// Assert
        //using var sheet = SpreadsheetAssert.SingleSheet(stream);
        //Assert.Equal("OLA", sheet["A1"].StringValue);
        //Assert.Equal("NORDMANN", sheet["C1"].StringValue);
    }


    //
    // CellStyleAttribute
    //


    //
    // CellValueConverterAttribute
    //


    //
    // CellValueTruncateAttribute
    //


    //
    // ColumnHeaderAttribute
    //


    //
    // ColumnOrderAttribute
    //


    //
    // ColumnWidthAttribute
    //
}
