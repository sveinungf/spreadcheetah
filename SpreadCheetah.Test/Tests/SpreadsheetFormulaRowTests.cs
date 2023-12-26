using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadCheetah.Styling;
using SpreadCheetah.Test.Helpers;
using System.Globalization;
using Xunit;
using Color = System.Drawing.Color;
using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;

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
        Assert.True(actualCell.CachedValue.IsBlank);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_AddRow_CellWithFormulaAndStyle(bool bold)
    {
        // Arrange
        const string formulaText = "SUM(A1,A2)";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        var style = new Style();
        style.Font.Bold = bold;
        var styleId = spreadsheet.AddStyle(style);

        var formula = new Formula(formulaText);
        var cell = new Cell(formula, styleId);

        // Act
        await spreadsheet.AddRowAsync(cell);
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(formulaText, actualCell.FormulaA1);
        Assert.Equal(bold, actualCell.Style.Font.Bold);
        Assert.True(actualCell.CachedValue.IsBlank);
    }

    [Theory]
    [MemberData(nameof(TestData.ValueTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_CellWithFormulaAndCachedValue(CellValueType valueType, RowCollectionType rowType, bool isNull)
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
            await spreadsheet.AddRowAsync(cell, rowType);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(formulaText, actualCell.FormulaA1);
        Assert.Equal(cachedValue.GetExpectedCachedValueAsString(), actualCell.CachedValue.ToString(CultureInfo.InvariantCulture));
        Assert.Equal(valueType.GetExpectedDefaultNumberFormat() ?? "", actualCell.Style.NumberFormat.Format);
    }

    public static IEnumerable<object?[]> ItalicWithValueTypes() => TestData.CombineWithValueTypes(true, false);

    [Theory]
    [MemberData(nameof(ItalicWithValueTypes))]
    public async Task Spreadsheet_AddRow_CellWithFormulaAndStyleAndCachedValue(bool italic, CellValueType cachedValueType, RowCollectionType rowType, bool cachedValueIsNull)
    {
        // Arrange
        const string formulaText = "SUM(A1,A2)";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");

        var style = new Style();
        style.Font.Italic = italic;
        var styleId = spreadsheet.AddStyle(style);

        var formula = new Formula(formulaText);
        var cell = CellFactory.Create(formula, cachedValueType, cachedValueIsNull, styleId, out var cachedValue);

        // Act
        await spreadsheet.AddRowAsync(cell, rowType);
        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(formulaText, actualCell.FormulaA1);
        Assert.Equal(italic, actualCell.Style.Font.Italic);
        Assert.Equal(cachedValue.GetExpectedCachedValueAsString(), actualCell.CachedValue.ToString(CultureInfo.InvariantCulture));
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
        Assert.True(actualCell.CachedValue.IsBlank);
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

    public static IEnumerable<object?[]> FormulaLengthsWithValueTypes() => TestData.CombineWithValueTypes(
        100,
        511,
        512,
        513,
        4100,
        8192);

    [Theory]
    [MemberData(nameof(FormulaLengthsWithValueTypes))]
    public async Task Spreadsheet_AddRow_CellWithVeryLongFormulaAndCachedValue(int formulaLength, CellValueType cachedValueType, RowCollectionType rowType, bool cachedValueIsNull)
    {
        // Arrange
        var formulaText = FormulaGenerator.Generate(formulaLength);
        object? cachedValue;
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");
            var formula = new Formula(formulaText);
            var cell = CellFactory.Create(formula, cachedValueType, cachedValueIsNull, null, out cachedValue);

            // Act
            await spreadsheet.AddRowAsync(cell, rowType);
            await spreadsheet.FinishAsync();
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(formulaText, actualCell.FormulaA1);
        Assert.Equal(cachedValue.GetExpectedCachedValueAsString(), actualCell.CachedValue.ToString(CultureInfo.InvariantCulture));
        Assert.Equal(cachedValueType.GetExpectedDefaultNumberFormat() ?? "", actualCell.Style.NumberFormat.Format);
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
        Assert.Equal(color.ToArgb(), actualCell.Style.Fill.BackgroundColor.Color.ToArgb());
        Assert.True(actualCell.CachedValue.IsBlank);
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
        var numberFormat = NumberFormat.Standard(StandardNumberFormat.Percent);
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options))
        {
            await spreadsheet.StartWorksheetAsync("Sheet");

            var style = new Style { Format = numberFormat };
            var styleId = spreadsheet.AddStyle(style);

            var formula = new Formula(formulaText);
            var cell = new Cell(formula, cachedValue, styleId);

            // Act
            await spreadsheet.AddRowAsync(cell);
            await spreadsheet.FinishAsync();
        }

        const int expectedNumberFormatId = (int)StandardNumberFormat.Percent;

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(formulaText, actualCell.FormulaA1);
        Assert.Equal(expectedNumberFormatId, actualCell.Style.NumberFormat.NumberFormatId);
        Assert.Equal(cachedValue, actualCell.CachedValue.GetNumber(), 0.01);
    }

    [Theory]
    [MemberData(nameof(TestData.ValueTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_ExplicitCellReferencesForFormulaCells(CellValueType valueType, RowCollectionType rowType, bool isNull)
    {
        // Arrange
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { DefaultDateTimeFormat = null, WriteCellReferenceAttributes = true };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        var formula = new Formula("SUM(A1,A2)");

        var row1 = Enumerable.Range(1, 10).Select(_ => CellFactory.Create(formula, valueType, isNull, null, out var value)).ToList();
        var row2 = Enumerable.Range(1, 1).Select(_ => CellFactory.Create(formula, valueType, isNull, null, out var value)).ToList();
        var row3 = Enumerable.Range(1, 100).Select(_ => CellFactory.Create(formula, valueType, isNull, null, out var value)).ToList();

        var expectedRow1Refs = Enumerable.Range(1, 10).Select(x => SpreadsheetUtility.GetColumnName(x) + "1").OfType<string?>();
        var expectedRow2Refs = Enumerable.Range(1, 1).Select(x => SpreadsheetUtility.GetColumnName(x) + "2").OfType<string?>();
        var expectedRow3Refs = Enumerable.Range(1, 100).Select(x => SpreadsheetUtility.GetColumnName(x) + "3").OfType<string?>();

        // Act
        await spreadsheet.StartWorksheetAsync("Sheet1");
        await spreadsheet.AddRowAsync(row1, rowType);
        await spreadsheet.AddRowAsync(row2, rowType);
        await spreadsheet.AddRowAsync(row3, rowType);

        await spreadsheet.StartWorksheetAsync("Sheet2");
        await spreadsheet.AddRowAsync(row1, rowType);

        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetParts = actual.WorkbookPart!.WorksheetParts.ToList();
        Assert.Equal(2, sheetParts.Count);

        var sheet1Rows = sheetParts[0].Worksheet.Descendants<Row>().ToList();
        Assert.Equal(3, sheet1Rows.Count);

        var actualRow1Refs = sheet1Rows[0].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        var actualRow2Refs = sheet1Rows[1].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        var actualRow3Refs = sheet1Rows[2].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);

        Assert.Equal(expectedRow1Refs, actualRow1Refs);
        Assert.Equal(expectedRow2Refs, actualRow2Refs);
        Assert.Equal(expectedRow3Refs, actualRow3Refs);

        var sheet2Row = sheetParts[1].Worksheet.Descendants<Row>().Single();
        var actualSheet2Refs = sheet2Row.Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        Assert.Equal(expectedRow1Refs, actualSheet2Refs);
    }

    [Theory]
    [MemberData(nameof(TestData.ValueTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_ExplicitCellReferencesForStyledFormulaCells(CellValueType valueType, RowCollectionType rowType, bool isNull)
    {
        // Arrange
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { WriteCellReferenceAttributes = true };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        var formula = new Formula("SUM(A1,A2)");
        var style = new Style();
        style.Fill.Color = Color.SaddleBrown;
        var styleId = spreadsheet.AddStyle(style);

        var row1 = Enumerable.Range(1, 10).Select(_ => CellFactory.Create(formula, valueType, isNull, styleId, out var value)).ToList();
        var row2 = Enumerable.Range(1, 1).Select(_ => CellFactory.Create(formula, valueType, isNull, styleId, out var value)).ToList();
        var row3 = Enumerable.Range(1, 100).Select(_ => CellFactory.Create(formula, valueType, isNull, styleId, out var value)).ToList();

        var expectedRow1Refs = Enumerable.Range(1, 10).Select(x => SpreadsheetUtility.GetColumnName(x) + "1").OfType<string?>();
        var expectedRow2Refs = Enumerable.Range(1, 1).Select(x => SpreadsheetUtility.GetColumnName(x) + "2").OfType<string?>();
        var expectedRow3Refs = Enumerable.Range(1, 100).Select(x => SpreadsheetUtility.GetColumnName(x) + "3").OfType<string?>();

        // Act
        await spreadsheet.StartWorksheetAsync("Sheet1");
        await spreadsheet.AddRowAsync(row1, rowType);
        await spreadsheet.AddRowAsync(row2, rowType);
        await spreadsheet.AddRowAsync(row3, rowType);

        await spreadsheet.StartWorksheetAsync("Sheet2");
        await spreadsheet.AddRowAsync(row1, rowType);

        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetParts = actual.WorkbookPart!.WorksheetParts.ToList();
        Assert.Equal(2, sheetParts.Count);

        var sheet1Rows = sheetParts[0].Worksheet.Descendants<Row>().ToList();
        Assert.Equal(3, sheet1Rows.Count);

        var actualRow1Refs = sheet1Rows[0].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        var actualRow2Refs = sheet1Rows[1].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        var actualRow3Refs = sheet1Rows[2].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);

        Assert.Equal(expectedRow1Refs, actualRow1Refs);
        Assert.Equal(expectedRow2Refs, actualRow2Refs);
        Assert.Equal(expectedRow3Refs, actualRow3Refs);

        var sheet2Row = sheetParts[1].Worksheet.Descendants<Row>().Single();
        var actualSheet2Refs = sheet2Row.Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        Assert.Equal(expectedRow1Refs, actualSheet2Refs);
    }

    [Theory]
    [MemberData(nameof(TestData.ValueTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_ExplicitCellReferencesForLongFormulaCells(CellValueType valueType, RowCollectionType rowType, bool isNull)
    {
        // Arrange
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize, WriteCellReferenceAttributes = true };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        var formula = new Formula(FormulaGenerator.Generate(options.BufferSize * 2));

        var row1 = Enumerable.Range(1, 10).Select(_ => CellFactory.Create(formula, valueType, isNull, null, out var value)).ToList();
        var row2 = Enumerable.Range(1, 1).Select(_ => CellFactory.Create(formula, valueType, isNull, null, out var value)).ToList();
        var row3 = Enumerable.Range(1, 100).Select(_ => CellFactory.Create(formula, valueType, isNull, null, out var value)).ToList();

        var expectedRow1Refs = Enumerable.Range(1, 10).Select(x => SpreadsheetUtility.GetColumnName(x) + "1").OfType<string?>();
        var expectedRow2Refs = Enumerable.Range(1, 1).Select(x => SpreadsheetUtility.GetColumnName(x) + "2").OfType<string?>();
        var expectedRow3Refs = Enumerable.Range(1, 100).Select(x => SpreadsheetUtility.GetColumnName(x) + "3").OfType<string?>();

        // Act
        await spreadsheet.StartWorksheetAsync("Sheet1");
        await spreadsheet.AddRowAsync(row1, rowType);
        await spreadsheet.AddRowAsync(row2, rowType);
        await spreadsheet.AddRowAsync(row3, rowType);

        await spreadsheet.StartWorksheetAsync("Sheet2");
        await spreadsheet.AddRowAsync(row1, rowType);

        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetParts = actual.WorkbookPart!.WorksheetParts.ToList();
        Assert.Equal(2, sheetParts.Count);

        var sheet1Rows = sheetParts[0].Worksheet.Descendants<Row>().ToList();
        Assert.Equal(3, sheet1Rows.Count);

        var actualRow1Refs = sheet1Rows[0].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        var actualRow2Refs = sheet1Rows[1].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        var actualRow3Refs = sheet1Rows[2].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);

        Assert.Equal(expectedRow1Refs, actualRow1Refs);
        Assert.Equal(expectedRow2Refs, actualRow2Refs);
        Assert.Equal(expectedRow3Refs, actualRow3Refs);

        var sheet2Row = sheetParts[1].Worksheet.Descendants<Row>().Single();
        var actualSheet2Refs = sheet2Row.Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        Assert.Equal(expectedRow1Refs, actualSheet2Refs);
    }

    [Theory]
    [MemberData(nameof(TestData.ValueTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_ExplicitCellReferencesForLongFormulaStyledCells(CellValueType valueType, RowCollectionType rowType, bool isNull)
    {
        // Arrange
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize, WriteCellReferenceAttributes = true };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options);
        var formula = new Formula(FormulaGenerator.Generate(options.BufferSize * 2));
        var style = new Style();
        style.Fill.Color = Color.LightGoldenrodYellow;
        var styleId = spreadsheet.AddStyle(style);

        var row1 = Enumerable.Range(1, 10).Select(_ => CellFactory.Create(formula, valueType, isNull, styleId, out var value)).ToList();
        var row2 = Enumerable.Range(1, 1).Select(_ => CellFactory.Create(formula, valueType, isNull, styleId, out var value)).ToList();
        var row3 = Enumerable.Range(1, 100).Select(_ => CellFactory.Create(formula, valueType, isNull, styleId, out var value)).ToList();

        var expectedRow1Refs = Enumerable.Range(1, 10).Select(x => SpreadsheetUtility.GetColumnName(x) + "1").OfType<string?>();
        var expectedRow2Refs = Enumerable.Range(1, 1).Select(x => SpreadsheetUtility.GetColumnName(x) + "2").OfType<string?>();
        var expectedRow3Refs = Enumerable.Range(1, 100).Select(x => SpreadsheetUtility.GetColumnName(x) + "3").OfType<string?>();

        // Act
        await spreadsheet.StartWorksheetAsync("Sheet1");
        await spreadsheet.AddRowAsync(row1, rowType);
        await spreadsheet.AddRowAsync(row2, rowType);
        await spreadsheet.AddRowAsync(row3, rowType);

        await spreadsheet.StartWorksheetAsync("Sheet2");
        await spreadsheet.AddRowAsync(row1, rowType);

        await spreadsheet.FinishAsync();

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetParts = actual.WorkbookPart!.WorksheetParts.ToList();
        Assert.Equal(2, sheetParts.Count);

        var sheet1Rows = sheetParts[0].Worksheet.Descendants<Row>().ToList();
        Assert.Equal(3, sheet1Rows.Count);

        var actualRow1Refs = sheet1Rows[0].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        var actualRow2Refs = sheet1Rows[1].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        var actualRow3Refs = sheet1Rows[2].Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);

        Assert.Equal(expectedRow1Refs, actualRow1Refs);
        Assert.Equal(expectedRow2Refs, actualRow2Refs);
        Assert.Equal(expectedRow3Refs, actualRow3Refs);

        var sheet2Row = sheetParts[1].Worksheet.Descendants<Row>().Single();
        var actualSheet2Refs = sheet2Row.Descendants<OpenXmlCell>().Select(x => x.CellReference?.Value);
        Assert.Equal(expectedRow1Refs, actualSheet2Refs);
    }
}
