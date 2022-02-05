using ClosedXML.Excel;
using SpreadCheetah.Test.Helpers;
using SpreadCheetah.Validations;
using Xunit;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetDataValidationTests
{
    [Theory]
    [InlineData("eq", XLOperator.EqualTo)]
    [InlineData("neq", XLOperator.NotEqualTo)]
    [InlineData("gt", XLOperator.GreaterThan)]
    [InlineData("gte", XLOperator.EqualOrGreaterThan)]
    [InlineData("lt", XLOperator.LessThan)]
    [InlineData("lte", XLOperator.EqualOrLessThan)]
    public async Task Spreadsheet_AddDataValidation_IntegerValue(string op, XLOperator expectedOperator)
    {
        // Arrange
        const int value = 100;
        var validation = op switch
        {
            "eq" => DataValidation.IntegerEqualTo(value),
            "neq" => DataValidation.IntegerNotEqualTo(value),
            "gt" => DataValidation.IntegerGreaterThan(value),
            "gte" => DataValidation.IntegerGreaterThanOrEqualTo(value),
            "lt" => DataValidation.IntegerLessThan(value),
            "lte" => DataValidation.IntegerLessThanOrEqualTo(value),
            _ => throw new ArgumentOutOfRangeException(nameof(op))
        };

        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            // Act
            spreadsheet.AddDataValidation("A1", validation);
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
        Assert.Equal(expectedOperator, actualValidation.Operator);
        Assert.Equal(value.ToString(), actualValidation.MinValue);
        Assert.Empty(actualValidation.MaxValue);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_AddDataValidation_IntegerTwoValues(bool between)
    {
        // Arrange
        const int min = -100;
        const int max = 200;
        var validation = between
            ? DataValidation.IntegerBetween(min, max)
            : DataValidation.IntegerNotBetween(min, max);
        var expectedOperator = between ? XLOperator.Between : XLOperator.NotBetween;

        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            // Act
            spreadsheet.AddDataValidation("A1", validation);
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
        Assert.Equal(expectedOperator, actualValidation.Operator);
        Assert.Equal(min.ToString(), actualValidation.MinValue);
        Assert.Equal(max.ToString(), actualValidation.MaxValue);
    }

    // TODO: Test for error/input
}
