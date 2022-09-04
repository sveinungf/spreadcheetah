using ClosedXML.Excel;
using SpreadCheetah.Styling;
using SpreadCheetah.Test.Helpers;
using System.Drawing;
using System.Globalization;
using Xunit;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetFormulaRowTests
{
    [Theory]
    [InlineData("1 + 2")]
    [InlineData("A1")]
    [InlineData("$A$1")]
    [InlineData("D17<>5")]
    [InlineData("SUM(A1,A2)")]
    [InlineData("SUM(A2:A8)/20")]
    [InlineData("IF(C2<D3, TRUE, FALSE)")]
    [InlineData("IF(C2<D3, \"Yes\", \"No\")")]
    public async Task Spreadsheet_AddRow_CellWithFormula(string? formulaText)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");
            var formula = new Formula(formulaText);
            var cell = new Cell(formula);

            // Act
            await spreadsheet.AddRowAsync(cell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(formulaText, actualCell.FormulaA1);
        Assert.Null(actualCell.CachedValue);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_AddRow_CellWithFormulaAndStyle(bool bold)
    {
        // Arrange
        const string formulaText = "SUM(A1,A2)";
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Font.Bold = bold;
            var styleId = spreadsheet.AddStyle(style);

            var formula = new Formula(formulaText);
            var cell = new Cell(formula, styleId);

            // Act
            await spreadsheet.AddRowAsync(cell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(formulaText, actualCell.FormulaA1);
        Assert.Equal(bold, actualCell.Style.Font.Bold);
        Assert.Null(actualCell.CachedValue);
    }

    [Theory]
    [MemberData(nameof(TestData.ValueTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_CellWithFormulaAndCachedValue(CellValueType valueType, bool isNull)
    {
        // Arrange
        const string formulaText = "SUM(A1,A2)";
        object? cachedValue;
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");
            var formula = new Formula(formulaText);
            var cell = CellFactory.Create(formula, valueType, isNull, null, out cachedValue);

            // Act
            await spreadsheet.AddRowAsync(cell);
            await spreadsheet.FinishAsync();
        }

        var expectedCachedValue = !isNull && cachedValue is DateTime dateTime
            ? dateTime.ToOADate().ToString()
            : Convert.ToString(cachedValue, CultureInfo.InvariantCulture);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(formulaText, actualCell.FormulaA1);
        Assert.Equal(expectedCachedValue, Convert.ToString(actualCell.CachedValue, CultureInfo.InvariantCulture));
        Assert.Equal(valueType.GetExpectedDefaultNumberFormat() ?? "", actualCell.Style.NumberFormat.Format);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_AddRow_CellWithFormulaAndStyleAndCachedValue(bool italic)
    {
        // Arrange
        const string formulaText = "SUM(A1,A2)";
        const int cachedValue = int.MinValue;
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Font.Italic = italic;
            var styleId = spreadsheet.AddStyle(style);

            var formula = new Formula(formulaText);
            var cell = new Cell(formula, cachedValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(cell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(formulaText, actualCell.FormulaA1);
        Assert.Equal(italic, actualCell.Style.Font.Italic);
        Assert.Equal(cachedValue.ToString(), actualCell.CachedValue);
    }

    [Theory]
    [InlineData(100)]
    [InlineData(511)]
    [InlineData(512)]
    [InlineData(513)]
    [InlineData(4100)]
    [InlineData(8192)]
    public async Task Spreadsheet_AddRow_CellWithVeryLongFormula(int length)
    {
        // Arrange
        var formulaText = FormulaGenerator.Generate(length);
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var formula = new Formula(formulaText);
            var cell = new Cell(formula);

            // Act
            await spreadsheet.AddRowAsync(cell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(formulaText, actualCell.FormulaA1);
        Assert.Null(actualCell.CachedValue);
    }

    [Theory]
    [InlineData(100)]
    [InlineData(511)]
    [InlineData(512)]
    [InlineData(513)]
    [InlineData(4100)]
    [InlineData(8192)]
    public async Task Spreadsheet_AddRow_CellWithVeryLongFormulaAndCachedStringValue(int length)
    {
        // Arrange
        var formulaText = FormulaGenerator.Generate(length);
        var cachedValue = new string('c', length);
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var formula = new Formula(formulaText);
            var cell = new Cell(formula, cachedValue);

            // Act
            await spreadsheet.AddRowAsync(cell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(formulaText, actualCell.FormulaA1);
        Assert.Equal(cachedValue, actualCell.CachedValue);
    }

    [Theory]
    [InlineData(100)]
    [InlineData(511)]
    [InlineData(512)]
    [InlineData(513)]
    [InlineData(4100)]
    [InlineData(8192)]
    public async Task Spreadsheet_AddRow_CellWithVeryLongFormulaAndCachedEmptyStringValue(int length)
    {
        // Arrange
        var formulaText = FormulaGenerator.Generate(length);
        const string cachedValue = "";
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var formula = new Formula(formulaText);
            var cell = new Cell(formula, cachedValue);

            // Act
            await spreadsheet.AddRowAsync(cell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(formulaText, actualCell.FormulaA1);
        Assert.Equal(cachedValue, actualCell.CachedValue);
    }

    [Theory]
    [InlineData(100)]
    [InlineData(511)]
    [InlineData(512)]
    [InlineData(513)]
    [InlineData(4100)]
    [InlineData(8192)]
    public async Task Spreadsheet_AddRow_CellWithVeryLongFormulaAndCachedDoubleValue(int length)
    {
        // Arrange
        var formulaText = FormulaGenerator.Generate(length);
        const double cachedValue = double.MaxValue;
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var formula = new Formula(formulaText);
            var cell = new Cell(formula, cachedValue);

            // Act
            await spreadsheet.AddRowAsync(cell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(formulaText, actualCell.FormulaA1);
        Assert.Equal(cachedValue.ToString(CultureInfo.InvariantCulture), actualCell.CachedValue);
    }

    [Theory]
    [InlineData(100, true)]
    [InlineData(100, false)]
    [InlineData(100, null)]
    [InlineData(8192, true)]
    [InlineData(8192, false)]
    [InlineData(8192, null)]
    public async Task Spreadsheet_AddRow_CellWithVeryLongFormulaAndCachedBooleanValue(int length, bool? cachedValue)
    {
        // Arrange
        var formulaText = FormulaGenerator.Generate(length);
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var formula = new Formula(formulaText);
            var cell = new Cell(formula, cachedValue);

            // Act
            await spreadsheet.AddRowAsync(cell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(formulaText, actualCell.FormulaA1);
        Assert.Equal(cachedValue, actualCell.CachedValue);
    }

    [Theory]
    [InlineData(100)]
    [InlineData(511)]
    [InlineData(512)]
    [InlineData(513)]
    [InlineData(4100)]
    [InlineData(8192)]
    public async Task Spreadsheet_AddRow_CellWithVeryLongFormulaAndStyle(int length)
    {
        // Arrange
        var formulaText = FormulaGenerator.Generate(length);
        var color = Color.Navy;
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style();
            style.Fill.Color = color;
            var styleId = spreadsheet.AddStyle(style);

            var formula = new Formula(formulaText);
            var cell = new Cell(formula, styleId);

            // Act
            await spreadsheet.AddRowAsync(cell);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(formulaText, actualCell.FormulaA1);
        Assert.Equal(color, actualCell.Style.Fill.BackgroundColor.Color);
        Assert.Null(actualCell.CachedValue);
    }

    [Theory]
    [InlineData(100)]
    [InlineData(511)]
    [InlineData(512)]
    [InlineData(513)]
    [InlineData(4100)]
    [InlineData(8192)]
    public async Task Spreadsheet_AddRow_CellWithVeryLongFormulaAndStyleAndCachedValue(int length)
    {
        // Arrange
        var formulaText = FormulaGenerator.Generate(length);
        const float cachedValue = 5.67f;
        const string numberFormat = NumberFormats.Percent;
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style { NumberFormat = numberFormat };
            var styleId = spreadsheet.AddStyle(style);

            var formula = new Formula(formulaText);
            var cell = new Cell(formula, cachedValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(cell);
            await spreadsheet.FinishAsync();
        }

        var expectedNumberFormatId = NumberFormats.GetPredefinedNumberFormatId(numberFormat) ?? 0;

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(formulaText, actualCell.FormulaA1);
        Assert.Equal(expectedNumberFormatId, actualCell.Style.NumberFormat.NumberFormatId);
        Assert.Equal(cachedValue.ToString(CultureInfo.InvariantCulture), actualCell.CachedValue);
    }
}
