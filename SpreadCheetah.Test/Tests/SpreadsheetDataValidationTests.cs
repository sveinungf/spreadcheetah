using ClosedXML.Excel;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using SpreadCheetah.Test.Helpers;
using SpreadCheetah.Validations;
using System.Globalization;
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
    public async Task Spreadsheet_AddDataValidation_DecimalValue(string op, XLOperator expectedOperator)
    {
        // Arrange
        const double value = 123.45;
        var validation = op switch
        {
            "eq" => DataValidation.DecimalEqualTo(value),
            "neq" => DataValidation.DecimalNotEqualTo(value),
            "gt" => DataValidation.DecimalGreaterThan(value),
            "gte" => DataValidation.DecimalGreaterThanOrEqualTo(value),
            "lt" => DataValidation.DecimalLessThan(value),
            "lte" => DataValidation.DecimalLessThanOrEqualTo(value),
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
        Assert.Equal(XLAllowedValues.Decimal, actualValidation.AllowedValues);
        Assert.Equal(expectedOperator, actualValidation.Operator);
        Assert.Equal(value.ToString(CultureInfo.InvariantCulture), actualValidation.MinValue);
        Assert.Empty(actualValidation.MaxValue);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_AddDataValidation_DecimalTwoValues(bool between)
    {
        // Arrange
        const double min = -44.55;
        const double max = 200.99;
        var validation = between
            ? DataValidation.DecimalBetween(min, max)
            : DataValidation.DecimalNotBetween(min, max);
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
        Assert.Equal(XLAllowedValues.Decimal, actualValidation.AllowedValues);
        Assert.Equal(expectedOperator, actualValidation.Operator);
        Assert.Equal(min.ToString(CultureInfo.InvariantCulture), actualValidation.MinValue);
        Assert.Equal(max.ToString(CultureInfo.InvariantCulture), actualValidation.MaxValue);
    }

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

    [Theory]
    [InlineData("eq", XLOperator.EqualTo)]
    [InlineData("neq", XLOperator.NotEqualTo)]
    [InlineData("gt", XLOperator.GreaterThan)]
    [InlineData("gte", XLOperator.EqualOrGreaterThan)]
    [InlineData("lt", XLOperator.LessThan)]
    [InlineData("lte", XLOperator.EqualOrLessThan)]
    public async Task Spreadsheet_AddDataValidation_TextLengthValue(string op, XLOperator expectedOperator)
    {
        // Arrange
        const int value = 10;
        var validation = op switch
        {
            "eq" => DataValidation.TextLengthEqualTo(value),
            "neq" => DataValidation.TextLengthNotEqualTo(value),
            "gt" => DataValidation.TextLengthGreaterThan(value),
            "gte" => DataValidation.TextLengthGreaterThanOrEqualTo(value),
            "lt" => DataValidation.TextLengthLessThan(value),
            "lte" => DataValidation.TextLengthLessThanOrEqualTo(value),
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
        Assert.Equal(XLAllowedValues.TextLength, actualValidation.AllowedValues);
        Assert.Equal(expectedOperator, actualValidation.Operator);
        Assert.Equal(value.ToString(), actualValidation.MinValue);
        Assert.Empty(actualValidation.MaxValue);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_AddDataValidation_TextLengthTwoValues(bool between)
    {
        // Arrange
        const int min = -100;
        const int max = 200;
        var validation = between
            ? DataValidation.TextLengthBetween(min, max)
            : DataValidation.TextLengthNotBetween(min, max);
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
        Assert.Equal(XLAllowedValues.TextLength, actualValidation.AllowedValues);
        Assert.Equal(expectedOperator, actualValidation.Operator);
        Assert.Equal(min.ToString(), actualValidation.MinValue);
        Assert.Equal(max.ToString(), actualValidation.MaxValue);
    }

    [Fact]
    public async Task Spreadsheet_AddDataValidation_ListValues()
    {
        // Arrange
        const string address = "A1";
        var values = new[] { "One", "\"Two\"", "<Three>" };
        var validation = DataValidation.ListValues(values);

        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            // Act
            spreadsheet.AddDataValidation(address, validation);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets.Single();
        var actualValidation = Assert.Single(worksheet.DataValidations);
        Assert.Equal(address, actualValidation.Address.Address);
        Assert.Equal(eDataValidationType.List, actualValidation.ValidationType.Type);
        var actualListValidation = (IExcelDataValidationList)actualValidation;
        Assert.Equal(values, actualListValidation.Formula.Values);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_AddDataValidation_ListValuesDropdown(bool showDropdown)
    {
        // Arrange
        var values = new[] { "One", "Two", "Three" };
        var validation = DataValidation.ListValues(values, showDropdown);

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
        Assert.Equal(XLAllowedValues.List, actualValidation.AllowedValues);
        Assert.Equal(showDropdown, actualValidation.InCellDropdown);
    }

    [Fact]
    public async Task Spreadsheet_AddDataValidation_InputAndErrorMessages()
    {
        // Arrange
        const string inputMessage = "This <is> the \"input\" message";
        const string errorMessage = "This <is> the \"error\" message";
        const string errorTitle = "This <is> the \"error\" title";
        var validation = DataValidation.IntegerGreaterThan(0);
        validation.InputMessage = inputMessage;
        validation.ErrorMessage = errorMessage;
        validation.ErrorTitle = errorTitle;
        validation.ErrorType = ValidationErrorType.Warning;

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
        Assert.Equal(inputMessage, actualValidation.InputMessage);
        Assert.Equal(errorMessage, actualValidation.ErrorMessage);
        Assert.Equal(errorTitle, actualValidation.ErrorTitle);
        Assert.Equal(XLErrorStyle.Warning, actualValidation.ErrorStyle);
    }
}
