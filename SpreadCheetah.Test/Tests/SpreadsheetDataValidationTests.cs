using ClosedXML.Excel;
using SpreadCheetah.Test.Helpers;
using SpreadCheetah.Validations;
using Xunit;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetDataValidationTests
{
    [Fact]
    public async Task Spreadsheet_AddDataValidation_IntegerBetween()
    {
        // Arrange
        const int min = 100;
        const int max = 200;
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            // Act
            spreadsheet.AddDataValidation("A1", DataValidation.IntegerBetween(min, max));
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualValidation = Assert.Single(worksheet.DataValidations);
        var actualRange = Assert.Single(actualValidation.Ranges);
        var cell = Assert.Single(actualRange.Cells());
        Assert.Equal(1, cell.Address.ColumnNumber);
        Assert.Equal(1, cell.Address.RowNumber);
        Assert.Equal(XLAllowedValues.WholeNumber, actualValidation.AllowedValues);
        Assert.Equal(XLOperator.Between, actualValidation.Operator);
        Assert.Equal(min.ToString(), actualValidation.MinValue);
        Assert.Equal(max.ToString(), actualValidation.MaxValue);
    }

    // TODO: Test for error/input
}
