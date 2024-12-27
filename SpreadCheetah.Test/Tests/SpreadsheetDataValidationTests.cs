using ClosedXML.Excel;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
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
    public async Task Spreadsheet_AddDataValidation_DateTimeValue(string op, XLOperator expectedOperator)
    {
        // Arrange
        var year = 2000;
        var month = 1;
        var day  = 1;
        var hour = 1;
        var minute = 1;
        var second = 1;
        DateTime value = new DateTime(year,
                                      month,
                                      day,
                                      hour,
                                      minute,
                                      second,
                                      DateTimeKind.Unspecified);

        var validation = op switch
        {
            "eq" => DataValidation.DateTimeEqualTo(value),
            "neq" => DataValidation.DateTimeNotEqualTo(value),
            "gt" => DataValidation.DateTimeGreaterThan(value),
            "gte" => DataValidation.DateTimeGreaterThanOrEqualTo(value),
            "lt" => DataValidation.DateTimeLessThan(value),
            "lte" => DataValidation.DateTimeLessThanOrEqualTo(value),
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
        Assert.Equal(XLAllowedValues.Date, actualValidation.AllowedValues);
        Assert.Equal(expectedOperator, actualValidation.Operator);


        var epoch = new DateTime(1900,1,1,1,1,1,DateTimeKind.Unspecified);
        var elapsedDates = (value.AddDays(2) - epoch).Days;
        Assert.Equal(elapsedDates.ToString(CultureInfo.InvariantCulture), actualValidation.MinValue);
        Assert.Empty(actualValidation.MaxValue);

        
    }


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
        const string reference = "A1";
        var values = new[] { "One", "\"Two\"", "<Three>" };
        var validation = DataValidation.ListValues(values);

        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            // Act
            spreadsheet.AddDataValidation(reference, validation);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets.Single();
        var actualValidation = Assert.Single(worksheet.DataValidations);
        Assert.Equal(reference, actualValidation.Address.Address);
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

    [Theory]
    [InlineData("A1:C1")]
    [InlineData("$A1:$C1")]
    [InlineData("A$1:C$1")]
    [InlineData("$A$1:$C$1")]
    public async Task Spreadsheet_AddDataValidation_ListValuesFromCells(string cellRange)
    {
        // Arrange
        const string reference = "A2";
        var validation = DataValidation.ListValuesFromCells(cellRange);

        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");
            await spreadsheet.AddRowAsync(new DataCell[] { new("Apple"), new("Orange"), new("Pear") });

            // Act
            spreadsheet.AddDataValidation(reference, validation);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets.Single();
        var actualValidation = Assert.Single(worksheet.DataValidations);
        Assert.Equal(reference, actualValidation.Address.Address);
        Assert.Equal(eDataValidationType.List, actualValidation.ValidationType.Type);
        var actualListValidation = (IExcelDataValidationList)actualValidation;
        Assert.Equal(cellRange, actualListValidation.Formula.ExcelFormula);
    }

    [Theory]
    [InlineData("ValueSheet")]
    [InlineData("Value sheet")]
    [InlineData("Value'sheet")]
    [InlineData("Value<sheet")]
    public async Task Spreadsheet_AddDataValidation_ListValuesFromCellsInAnotherWorksheet(string valueSheetName)
    {
        // Arrange
        const string reference = "A2";
        const string cellRange = "A1:C1";
        var validation = DataValidation.ListValuesFromCells(valueSheetName, cellRange);

        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync(valueSheetName);
            await spreadsheet.AddRowAsync(new DataCell[] { new("Apple"), new("Orange"), new("Pear") });
            await spreadsheet.StartWorksheetAsync("Data validation sheet");

            // Act
            spreadsheet.AddDataValidation(reference, validation);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets.Single(x => !string.Equals(x.Name, valueSheetName, StringComparison.Ordinal));
        var actualValidation = Assert.Single(worksheet.DataValidations);
        Assert.Equal(reference, actualValidation.Address.Address);
        Assert.Equal(eDataValidationType.List, actualValidation.ValidationType.Type);
        var actualListValidation = (IExcelDataValidationList)actualValidation;
#pragma warning disable CA1307, MA0074 // Specify StringComparison for clarity
        var expectedValueSheetName = valueSheetName?.Replace("'", "''");
#pragma warning restore CA1307, MA0074 // Specify StringComparison for clarity
        Assert.Equal($"'{expectedValueSheetName}'!{cellRange}", actualListValidation.Formula.ExcelFormula);
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

    [Fact]
    public async Task Spreadsheet_AddDataValidation_CellRange()
    {
        // Arrange
        var validation = DataValidation.IntegerGreaterThan(0);

        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            // Act
            spreadsheet.AddDataValidation("A1:C10", validation);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualValidation = Assert.Single(worksheet.DataValidations);
        var actualRange = Assert.Single(actualValidation.Ranges);
        Assert.Equal(3, actualRange.ColumnCount());
        Assert.Equal(10, actualRange.RowCount());

        var firstCell = actualRange.Cells().First();
        Assert.Equal(1, firstCell.Address.ColumnNumber);
        Assert.Equal(1, firstCell.Address.RowNumber);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("A 1")]
    [InlineData("A1 ")]
    [InlineData(" A1")]
    [InlineData("a1")]
    [InlineData("Å1")]
    [InlineData("A１")]
    [InlineData("AAAA1")]
    [InlineData("A12345678")]
    [InlineData("A0")]
    [InlineData("$A$0")]
    [InlineData("A1:A0")]
    public async Task Spreadsheet_AddDataValidation_InvalidReference(string? reference)
    {
        // Arrange
        var validation = DataValidation.IntegerGreaterThan(0);

        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);

        await spreadsheet.StartWorksheetAsync("Sheet");

        // Act & Assert
        Assert.ThrowsAny<ArgumentException>(() => spreadsheet.AddDataValidation(reference!, validation));
    }

    [Fact]
    public async Task Spreadsheet_AddDataValidation_OverwritePrevious()
    {
        // Arrange
        var previousValidation = DataValidation.TextLengthLessThan(10);
        var validation = DataValidation.TextLengthLessThan(20);

        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");
            spreadsheet.AddDataValidation("A1", previousValidation);

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
        Assert.Equal(XLOperator.LessThan, actualValidation.Operator);
        Assert.Equal("20", actualValidation.MinValue);
    }

    [Fact]
    public async Task Spreadsheet_AddDataValidation_ThrowAfterMaxNumberOfValidations()
    {
        // Arrange
        var validation = DataValidation.IntegerBetween(0, 9);

        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        for (var i = 1; i <= 65534; ++i)
        {
            var reference = "A" + i;
            spreadsheet.AddDataValidation(reference, validation);
        }

        // Act & Assert
        Assert.ThrowsAny<SpreadCheetahException>(() => spreadsheet.AddDataValidation("B1", validation));
        await spreadsheet.FinishAsync();
        SpreadsheetAssert.Valid(stream);
    }

    [Fact]
    public async Task Spreadsheet_AddDataValidation_TwoWorksheetsCanHaveMaxNumberOfValidations()
    {
        // Arrange
        static void AddMaxNumberOfDataValidations(string column, Spreadsheet spreadsheet)
        {
            var validation = DataValidation.IntegerBetween(0, 9);
            for (var i = 1; i <= 65534; ++i)
            {
                var reference = column + i;
                spreadsheet.AddDataValidation(reference, validation);
            }
        }

        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);

        // Act
        await spreadsheet.StartWorksheetAsync("Sheet");
        AddMaxNumberOfDataValidations("A", spreadsheet);
        await spreadsheet.StartWorksheetAsync("Sheet 2");
        AddMaxNumberOfDataValidations("B", spreadsheet);
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
    }

    [Fact]
    public async Task Spreadsheet_AddDataValidation_Multiple()
    {
        // Arrange
        const int count = 1000;
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var dataValidations = DataValidationGenerator.Generate(count);

        for (var i = 0; i < dataValidations.Count; ++i)
        {
            var reference = "A" + (i + 1);

            // Act
            spreadsheet.AddDataValidation(reference, dataValidations[i]);
        }

        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualValidations = worksheet.Cells().Select(x => x.GetDataValidation()).ToList();
        Assert.All(dataValidations.Zip(actualValidations), x => SpreadsheetAssert.EquivalentDataValidation(x.First, x.Second));
    }
}
