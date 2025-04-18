using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using SpreadCheetah.Styling;
using SpreadCheetah.Test.Helpers;
using SpreadCheetah.Worksheets;
using System.Globalization;
using CellType = SpreadCheetah.Test.Helpers.CellType;
using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using OpenXmlCellValue = DocumentFormat.OpenXml.Spreadsheet.CellValues;
using SpreadsheetAssert = SpreadCheetah.TestHelpers.Assertions.SpreadsheetAssert;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetRowTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Spreadsheet_AddRow_ThrowsWhenNoWorksheet(bool hasWorksheet)
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);

        if (hasWorksheet)
            await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        // Act
        var exception = await Record.ExceptionAsync(async () => await spreadsheet.AddRowAsync([], Token));

        // Assert
        Assert.NotEqual(hasWorksheet, exception is SpreadCheetahException);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Spreadsheet_AddRow_ThrowsWhenAlreadyFinished(bool finished)
    {
        // Arrange
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        if (finished)
            await spreadsheet.FinishAsync(Token);

        // Act
        var exception = await Record.ExceptionAsync(async () => await spreadsheet.AddRowAsync([], Token));

        // Assert
        Assert.Equal(finished, exception != null);
    }

    [Theory]
    [MemberData(nameof(TestData.CellTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_EmptyRow(CellType type, RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

            var addRowTask = type switch
            {
                CellType.Cell => spreadsheet.AddRowAsync(Enumerable.Empty<Cell>(), rowType),
                CellType.DataCell => spreadsheet.AddRowAsync(Enumerable.Empty<DataCell>(), rowType),
                CellType.StyledCell => spreadsheet.AddRowAsync(Enumerable.Empty<StyledCell>(), rowType),
                _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
            };

            // Act
            await addRowTask;
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        Assert.Empty(sheetPart.Worksheet.Descendants<OpenXmlCell>());
    }

    [Theory]
    [MemberData(nameof(TestData.CellTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_CellWithoutValue(CellType type, RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
            var cell = CellFactory.CreateWithoutValue(type);

            // Act
            await spreadsheet.AddRowAsync(cell, rowType);
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var actualCell = sheetPart.Worksheet.Descendants<OpenXmlCell>().Single();
        Assert.Null(actualCell.DataType?.Value);
        Assert.Equal(string.Empty, actualCell.InnerText);
    }

    public static IEnumerable<object?[]> Strings() => TestData.CombineWithCellTypes(
        "OneWord",
        "With whitespace",
        "With trailing whitespace ",
        " With leading whitespace",
        "With-Special!#¤%Characters",
        "With-Special<>&'\\Characters",
        "With\"Quotation\"Marks",
        "WithNorwegianCharactersÆØÅ",
        "With\ud83d\udc4dEmoji",
        "With🌉Emoji",
        "With\tValid\r\nControlCharacters",
        "WithCharacters\u00a0\u00c9\u00ffBetween160And255",
        "",
        null);

    [Theory]
    [MemberData(nameof(Strings))]
    public async Task Spreadsheet_AddRow_CellWithStringValue(string? value, CellType type, RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var cell = CellFactory.Create(type, value);

        // Act
        await spreadsheet.AddRowAsync(cell, rowType);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(value, sheet["A1"].StringValue);
    }

    [Theory]
    [MemberData(nameof(Strings))]
    public async Task Spreadsheet_AddRow_CellWithReadOnlyMemoryOfCharValue(string? value, CellType type, RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var cell = CellFactory.Create(type, value.AsMemory());

        // Act
        await spreadsheet.AddRowAsync(cell, rowType);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(value ?? "", sheet["A1"].StringValue);
    }

    [Theory]
    [MemberData(nameof(TestData.CellTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_CellWithInvalidControlCharacterStringValue(CellType type, RowCollectionType rowType)
    {
        // Arrange
        const string value = "With\u0000Control\u0010\u001fCharacters";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var cell = CellFactory.Create(type, value);

        // Act
        await spreadsheet.AddRowAsync(cell, rowType);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var actualCell = sheetPart.Worksheet.Descendants<OpenXmlCell>().Single();
        Assert.Equal(OpenXmlCellValue.InlineString, actualCell.DataType?.Value);
        Assert.Equal("WithControlCharacters", actualCell.InnerText);
    }

    public static IEnumerable<object?[]> StringLengths() => TestData.CombineWithCellTypes(
        4095,
        4096,
        4097,
        10000,
        30000,
        32767);

    [Theory]
    [MemberData(nameof(StringLengths))]
    public async Task Spreadsheet_AddRow_CellWithVeryLongStringValue(int length, CellType type, RowCollectionType rowType)
    {
        // Arrange
        var value = new string('a', length);
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token))
        {
            await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
            var cell = CellFactory.Create(type, value);

            // Act
            await spreadsheet.AddRowAsync(cell, rowType);
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var actualCell = sheetPart.Worksheet.Descendants<OpenXmlCell>().Single();
        Assert.Equal(OpenXmlCellValue.InlineString, actualCell.DataType?.Value);
        Assert.Equal(value, actualCell.InnerText);
    }

    public static IEnumerable<object?[]> Integers() => TestData.CombineWithCellTypes<int?>(
        1234,
        0,
        -1234,
        int.MinValue,
        int.MaxValue,
        null);

    [Theory]
    [MemberData(nameof(Integers))]
    public async Task Spreadsheet_AddRow_CellWithIntegerValue(int? value, CellType type, RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
            var cell = CellFactory.Create(type, value);

            // Act
            await spreadsheet.AddRowAsync(cell, rowType);
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var actualCell = sheetPart.Worksheet.Descendants<OpenXmlCell>().Single();
        Assert.Equal(OpenXmlCellValue.Number, actualCell.GetDataType());
        Assert.Equal(value?.ToString() ?? string.Empty, actualCell.InnerText);
    }

    public static IEnumerable<object?[]> Longs() => TestData.CombineTuplesWithCellTypes<long?, string>(
        (1234L, "1234"),
        (0L, "0"),
        (-1234L, "-1234"),
        (314748364700000L, "314748364700000"),
#if NET472_OR_GREATER
        (long.MinValue, "-9.22337203685478E+18"),
        (long.MaxValue, "9.22337203685478E+18"),
#else
        (long.MinValue, "-9.223372036854776E+18"),
        (long.MaxValue, "9.223372036854776E+18"),
#endif
        (null, ""));

    [Theory]
    [MemberData(nameof(Longs))]
    public async Task Spreadsheet_AddRow_CellWithLongValue(long? initialValue, string expectedValue, CellType type, RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
            var cell = CellFactory.Create(type, initialValue);

            // Act
            await spreadsheet.AddRowAsync(cell, rowType);
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var actualCell = sheetPart.Worksheet.Descendants<OpenXmlCell>().Single();
        Assert.Equal(OpenXmlCellValue.Number, actualCell.GetDataType());
        Assert.Equal(expectedValue, actualCell.InnerText);
    }

    public static IEnumerable<object?[]> Floats() => TestData.CombineTuplesWithCellTypes<float?, string>(
        (1234f, "1234"),
        (0.1f, "0.1"),
        (0.0f, "0"),
        (-0.1f, "-0.1"),
        (0.1111111f, "0.1111111"),
        (11.11111f, "11.11111"),
        (2.222222E+38f, "2.222222E+38"),
        (-0.3333333f, "-0.3333333"),
        (null, ""));

    [Theory]
    [MemberData(nameof(Floats))]
    public async Task Spreadsheet_AddRow_CellWithFloatValue(float? initialValue, string expectedValue, CellType type, RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
            var cell = CellFactory.Create(type, initialValue);

            // Act
            await spreadsheet.AddRowAsync(cell, rowType);
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var actualCell = sheetPart.Worksheet.Descendants<OpenXmlCell>().Single();
        Assert.Equal(OpenXmlCellValue.Number, actualCell.GetDataType());
        Assert.Equal(expectedValue, actualCell.InnerText);
    }

    public static IEnumerable<object?[]> Doubles() => TestData.CombineTuplesWithCellTypes<double?, string>(
        (1234d, "1234"),
        (0.1, "0.1"),
        (0.0, "0"),
        (-0.1, "-0.1"),
        (0.1111111111111, "0.1111111111111"),
        (11.1111111111111, "11.1111111111111"),
#if NET472_OR_GREATER
        (11.11111111111111111111, "11.1111111111111"),
#else
        (11.11111111111111111111, "11.11111111111111"),
#endif
        (2.2222222222E+50, "2.2222222222E+50"),
        (-0.3333333, "-0.3333333"),
        (null, ""));

    [Theory]
    [MemberData(nameof(Doubles))]
    public async Task Spreadsheet_AddRow_CellWithDoubleValue(double? initialValue, string expectedValue, CellType type, RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
            var cell = CellFactory.Create(type, initialValue);

            // Act
            await spreadsheet.AddRowAsync(cell, rowType);
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var actualCell = sheetPart.Worksheet.Descendants<OpenXmlCell>().Single();
        Assert.Equal(OpenXmlCellValue.Number, actualCell.GetDataType());
        Assert.Equal(expectedValue, actualCell.InnerText);
    }

    public static IEnumerable<object?[]> Decimals() => TestData.CombineTuplesWithCellTypes(
        ("1234", "1234"),
        ("0.1", "0.1"),
        ("0.0", "0"),
        ("-0.1", "-0.1"),
        ("-0.3333333", "-0.3333333"),
        ("0.1111111111111", "0.1111111111111"),
        ("11.1111111111111", "11.1111111111111"),
#if NET472_OR_GREATER
        ("11.11111111111111111111", "11.1111111111111"),
        ("0.123456789012345678901234567", "0.123456789012346"),
#else
        ("11.11111111111111111111", "11.11111111111111"),
        ("0.123456789012345678901234567", "0.12345678901234568"),
#endif
        (null, ""));

    [Theory]
    [MemberData(nameof(Decimals))]
    public async Task Spreadsheet_AddRow_CellWithDecimalValue(string? initialValue, string expectedValue, CellType type, RowCollectionType rowType)
    {
        // Arrange
        var decimalValue = initialValue != null ? decimal.Parse(initialValue, CultureInfo.InvariantCulture) : null as decimal?;
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
            var cell = CellFactory.Create(type, decimalValue);

            // Act
            await spreadsheet.AddRowAsync(cell, rowType);
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var actualCell = sheetPart.Worksheet.Descendants<OpenXmlCell>().Single();
        Assert.Equal(OpenXmlCellValue.Number, actualCell.GetDataType());
        Assert.Equal(expectedValue, actualCell.InnerText);
    }

    public static IEnumerable<object?[]> DateTimes() => TestData.CombineWithCellTypes<DateTime?>(
        new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Unspecified),
        new DateTime(2001, 2, 3, 4, 5, 6, DateTimeKind.Unspecified),
        new DateTime(2001, 2, 3, 4, 5, 6, 789, DateTimeKind.Unspecified),
        DateTime.MaxValue,
        DateTime.MinValue,
        null
    );

    [Theory]
    [MemberData(nameof(DateTimes))]
    public async Task Spreadsheet_AddRow_CellWithDateTimeValue(DateTime? value, CellType type, RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
            var cell = CellFactory.Create(type, value);

            // Act
            await spreadsheet.AddRowAsync(cell, rowType);
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualCell = worksheet.Cell(1, 1);
        Assert.Equal(value?.ToOADate() ?? 0, actualCell.GetValue<double?>() ?? 0, 0.0001);
        Assert.Equal(NumberFormats.DateTimeSortable, actualCell.Style.NumberFormat.Format);
    }

#pragma warning disable CS0618 // Type or member is obsolete - Testing for backwards compatibilty
    public static IEnumerable<object?[]> DateTimeNumberFormats() => TestData.CombineWithCellTypes(
        NumberFormats.DateTimeSortable,
        NumberFormats.General,
        "mm-dd-yy",
        "yyyy",
        null
    );

    [Theory]
    [MemberData(nameof(DateTimeNumberFormats))]
    public async Task Spreadsheet_AddRow_CellWithDateTimeValueAndDefaultNumberFormat(string? defaultNumberFormat, CellType type, RowCollectionType rowType)
    {
        // Arrange
        var value = new DateTime(2022, 9, 11, 14, 7, 13, DateTimeKind.Unspecified);
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { DefaultDateTimeNumberFormat = defaultNumberFormat };
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token))
        {
            await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
            var cell = CellFactory.Create(type, value);

            // Act
            await spreadsheet.AddRowAsync(cell, rowType);
            await spreadsheet.FinishAsync(Token);
        }

        var expectedValue = DateTime.FromOADate(value.ToOADate());

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var package = new ExcelPackage(stream);
        var worksheet = package.Workbook.Worksheets.Single();
        var actualCell = worksheet.Cells.Single();
        Assert.Equal(expectedValue, actualCell.GetValue<DateTime>());
        Assert.Equal(defaultNumberFormat ?? NumberFormats.General, actualCell.Style.Numberformat.Format);
    }
#pragma warning restore CS0618 // Type or member is obsolete

    public static IEnumerable<object?[]> Booleans() => TestData.CombineTuplesWithCellTypes<bool?, string>(
        (true, "1"),
        (false, "0"),
        (null, ""));

    [Theory]
    [MemberData(nameof(Booleans))]
    public async Task Spreadsheet_AddRow_CellWithBooleanValue(bool? initialValue, string expectedValue, CellType type, RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
            var cell = CellFactory.Create(type, initialValue);

            // Act
            await spreadsheet.AddRowAsync(cell, rowType);
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetPart = actual.WorkbookPart!.WorksheetParts.Single();
        var actualCell = sheetPart.Worksheet.Descendants<OpenXmlCell>().Single();
        var expectedDataType = initialValue is null ? OpenXmlCellValue.Number : OpenXmlCellValue.Boolean;
        Assert.Equal(expectedDataType, actualCell.GetDataType());
        Assert.Equal(expectedValue, actualCell.InnerText);
    }

    [Theory]
    [MemberData(nameof(TestData.CellTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_MultipleColumns(CellType type, RowCollectionType rowType)
    {
        // Arrange
        var values = Enumerable.Range(1, 1000).Select(x => x.ToString()).ToList();
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
            var cells = values.ConvertAll(x => CellFactory.Create(type, x));

            // Act
            await spreadsheet.AddRowAsync(cells, rowType);
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualValues = worksheet.Cells(true).Select(x => x.Value.GetText()).ToList();
        Assert.Equal(values, actualValues);
    }

    [Theory]
    [MemberData(nameof(TestData.CellTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_MultipleColumnsWithVeryLongStringValues(CellType type, RowCollectionType rowType)
    {
        // Arrange
        var values = Enumerable.Range(1, 1000).Select(x => new string((char)('A' + (x % 26)), 1000)).ToList();
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var cells = values.ConvertAll(x => CellFactory.Create(type, x));

        // Act
        await spreadsheet.AddRowAsync(cells, rowType);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualValues = worksheet.Cells(true).Select(x => x.Value.GetText()).ToList();
        Assert.Equal(values, actualValues);
    }

    [Theory]
    [MemberData(nameof(TestData.CellTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_MultipleRows(CellType type, RowCollectionType rowType)
    {
        // Arrange
        var values = Enumerable.Range(1, 1000).Select(x => x.ToString()).ToList();
        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

            // Act
            foreach (var value in values)
            {
                await spreadsheet.AddRowAsync(CellFactory.Create(type, value), rowType);
            }

            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualValues = worksheet.Cells(true).Select(x => x.Value.GetText()).ToList();
        Assert.Equal(values, actualValues);
    }

    [Theory]
    [MemberData(nameof(TestData.CellTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_MultipleRowsWithCharactersToEscape(CellType type, RowCollectionType rowType)
    {
        // Arrange
        const string cellValue = "'This \" text & has & &&&&&&' characters < that > needs & to & be & escaped!'";
        var values = Enumerable.Repeat(cellValue, 1000);
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        // Act
        foreach (var value in values)
        {
            await spreadsheet.AddRowAsync(CellFactory.Create(type, value), rowType);
        }

        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualValues = worksheet.Cells(true).Select(x => x.Value.GetText());
        Assert.All(actualValues, x => Assert.Equal(cellValue, x));
    }

    [Theory]
    [MemberData(nameof(TestData.CellAndValueTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_ExplicitCellReferences(CellValueType valueType, bool isNull, CellType cellType, RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { DefaultDateTimeFormat = null, WriteCellReferenceAttributes = true };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);

        var row1 = Enumerable.Range(1, 10).Select(_ => CellFactory.Create(cellType, valueType, isNull, null, out var _)).ToList();
        var row2 = Enumerable.Range(1, 1).Select(_ => CellFactory.Create(cellType, valueType, isNull, null, out var _)).ToList();
        var row3 = Enumerable.Range(1, 100).Select(_ => CellFactory.Create(cellType, valueType, isNull, null, out var _)).ToList();

        var expectedRow1Refs = Enumerable.Range(1, 10).Select(x => SpreadsheetUtility.GetColumnName(x) + "1").OfType<string?>();
        var expectedRow2Refs = Enumerable.Range(1, 1).Select(x => SpreadsheetUtility.GetColumnName(x) + "2").OfType<string?>();
        var expectedRow3Refs = Enumerable.Range(1, 100).Select(x => SpreadsheetUtility.GetColumnName(x) + "3").OfType<string?>();

        // Act
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        await spreadsheet.AddRowAsync(row1, rowType);
        await spreadsheet.AddRowAsync(row2, rowType);
        await spreadsheet.AddRowAsync(row3, rowType);

        await spreadsheet.StartWorksheetAsync("Sheet 2", token: Token);
        await spreadsheet.AddRowAsync(row1, rowType);

        await spreadsheet.FinishAsync(Token);

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
    [MemberData(nameof(TestData.CellTypes), MemberType = typeof(TestData))]
    public async Task Spreadsheet_AddRow_ExplicitCellReferencesForLongStringValueCells(CellType cellType, RowCollectionType rowType)
    {
        // Arrange
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { BufferSize = SpreadCheetahOptions.MinimumBufferSize, WriteCellReferenceAttributes = true };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        var value = new string('a', options.BufferSize * 2);

        var row1 = Enumerable.Range(1, 10).Select(_ => CellFactory.Create(cellType, value)).ToList();
        var row2 = Enumerable.Range(1, 1).Select(_ => CellFactory.Create(cellType, value)).ToList();
        var row3 = Enumerable.Range(1, 100).Select(_ => CellFactory.Create(cellType, value)).ToList();

        var expectedRow1Refs = Enumerable.Range(1, 10).Select(x => SpreadsheetUtility.GetColumnName(x) + "1").OfType<string?>();
        var expectedRow2Refs = Enumerable.Range(1, 1).Select(x => SpreadsheetUtility.GetColumnName(x) + "2").OfType<string?>();
        var expectedRow3Refs = Enumerable.Range(1, 100).Select(x => SpreadsheetUtility.GetColumnName(x) + "3").OfType<string?>();

        // Act
        await spreadsheet.StartWorksheetAsync("Sheet 1", token: Token);
        await spreadsheet.AddRowAsync(row1, rowType);
        await spreadsheet.AddRowAsync(row2, rowType);
        await spreadsheet.AddRowAsync(row3, rowType);

        await spreadsheet.StartWorksheetAsync("Sheet 2", token: Token);
        await spreadsheet.AddRowAsync(row1, rowType);

        await spreadsheet.FinishAsync(Token);

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
    [InlineData(true)]
    [InlineData(false)]
    public async Task Spreadsheet_AddRow_ExplicitCellReferenceForDateTimeCellWithDefaultStyling(bool isNull)
    {
        // Arrange
        using var stream = new MemoryStream();
        var options = new SpreadCheetahOptions { WriteCellReferenceAttributes = true };
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, options, Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        DateTime? cellValue = isNull ? null : new DateTime(2019, 8, 7, 6, 5, 4, DateTimeKind.Utc);
        var cell = new Cell(cellValue);

        // Act
        await spreadsheet.AddRowAsync(cell);
        await spreadsheet.FinishAsync(Token);

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var actual = SpreadsheetDocument.Open(stream, true);
        var sheetPart = Assert.Single(actual.WorkbookPart!.WorksheetParts);
        var actualRow = Assert.Single(sheetPart.Worksheet.Descendants<Row>());
        var actualCell = Assert.Single(actualRow.Descendants<OpenXmlCell>());
        Assert.Equal("A1", actualCell.CellReference?.Value);
    }

    public static IEnumerable<object?[]> RowHeights() => TestData.CombineWithCellTypes<double?>(
        0.1,
        10d,
        123.456,
        409d,
        null);

    [Theory]
    [MemberData(nameof(RowHeights))]
    public async Task Spreadsheet_AddRow_RowHeight(double? height, CellType type, RowCollectionType rowType)
    {
        // Arrange
        var rowOptions = new RowOptions { Height = height };
        var expectedHeight = height ?? 15;

        using var stream = new MemoryStream();
        await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
        {
            await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
            await spreadsheet.AddRowAsync(CellFactory.Create(type, "My value"), rowType, rowOptions);
            await spreadsheet.FinishAsync(Token);
        }

        // Assert
        SpreadsheetAssert.Valid(stream);
        using var workbook = new XLWorkbook(stream);
        var worksheet = workbook.Worksheets.Single();
        var actualHeight = worksheet.Row(1).Height;
        Assert.Equal(expectedHeight, actualHeight, 5);
    }
}
