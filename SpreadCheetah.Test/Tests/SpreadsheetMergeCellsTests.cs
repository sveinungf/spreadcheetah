using ClosedXML.Excel;
using SpreadCheetah.Test.Helpers;
using SpreadCheetah.Validations;
using SpreadCheetah.Worksheets;
using Xunit;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetMergeCellsTests
{
    [Theory]
    [InlineData("A1:A2")]
    [InlineData("A1:F1")]
    [InlineData("A1:XY10000")]
    public async Task Spreadsheet_MergeCells_ValidRange(string cellRange)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        // Act
        spreadsheet.MergeCells(cellRange);
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualMergedRange = Assert.Single(worksheet.MergedRanges);
        Assert.Equal(cellRange, actualMergedRange.RangeAddress.ToString());
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("A")]
    [InlineData("A1")]
    [InlineData("A1:A")]
    [InlineData("A1:AAAA2")]
    [InlineData("$A$1:$A$2")]
    public async Task Spreadsheet_MergeCells_InvalidRange(string cellRange)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        // Act & Assert
        Assert.ThrowsAny<ArgumentException>(() => spreadsheet.MergeCells(cellRange));
    }

    [Fact]
    public async Task Spreadsheet_MergeCells_WorksheetWithAutoFilterAndDataValidation()
    {
        // Arrange
        const string autoFilterRange = "A1:F1";
        const string dataValidationRange = "A2:A100";
        const string mergeRange = "B2:F3";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        var options = new WorksheetOptions { AutoFilter = new AutoFilterOptions(autoFilterRange) };
        await spreadsheet.StartWorksheetAsync("Sheet", options);

        const int validationValue = 50;
        var validation = DataValidation.TextLengthLessThan(validationValue);
        spreadsheet.AddDataValidation(dataValidationRange, validation);

        // Act
        spreadsheet.MergeCells(mergeRange);
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualMergedRange = Assert.Single(worksheet.MergedRanges);
        Assert.Equal(mergeRange, actualMergedRange.RangeAddress.ToString());
        Assert.Equal(autoFilterRange, worksheet.AutoFilter.Range.RangeAddress.ToString());
        var actualValidation = Assert.Single(worksheet.DataValidations);
        var actualRange = Assert.Single(actualValidation.Ranges);
        Assert.Equal(dataValidationRange, actualRange.RangeAddress.ToString());
        Assert.Equal(XLAllowedValues.TextLength, actualValidation.AllowedValues);
        Assert.Equal(XLOperator.LessThan, actualValidation.Operator);
        Assert.Equal(validationValue.ToString(), actualValidation.MinValue);
        Assert.Empty(actualValidation.MaxValue);
    }
}
