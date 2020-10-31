using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadCheetah.Test.Helpers;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;

namespace SpreadCheetah.Test.Tests
{
    public class SpreadsheetRowTests
    {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task Spreadsheet_AddRow_ThrowsWhenAlreadyFinished(bool finished)
        {
            // Arrange
            using var spreadsheet = await Spreadsheet.CreateNewAsync(Stream.Null);
            await spreadsheet.StartWorksheetAsync("Sheet");

            if (finished)
                await spreadsheet.FinishAsync();

            // Act
            var exception = await Record.ExceptionAsync(async () => await spreadsheet.AddRowAsync(Array.Empty<DataCell>()));

            // Assert
            Assert.Equal(finished, exception != null);
        }

        [Theory]
        [InlineData("OneWord")]
        [InlineData("With whitespace")]
        [InlineData("With trailing whitespace ")]
        [InlineData(" With leading whitespace")]
        [InlineData("With-Special-Characters!#¤%&")]
        [InlineData("With\"Quotation\"Marks")]
        [InlineData("WithNorwegianCharactersÆØÅ")]
        [InlineData("WithEmoji\ud83d\udc4d")]
        [InlineData("")]
        [InlineData(null)]
        public async Task Spreadsheet_AddRow_CellWithStringValue(string? value)
        {
            // Arrange
            using var stream = new MemoryStream();
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");
                var cell = new DataCell(value);

                // Act
                await spreadsheet.AddRowAsync(cell);
                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var actual = SpreadsheetDocument.Open(stream, true);
            var sheetPart = actual.WorkbookPart.WorksheetParts.Single();
            var actualCell = sheetPart.Worksheet.Descendants<Cell>().Single();
            Assert.Equal(CellValues.InlineString, actualCell.DataType.Value);
            Assert.Equal(value ?? string.Empty, actualCell.InnerText);
        }

        [Theory]
        [InlineData(4095)]
        [InlineData(4096)]
        [InlineData(4097)]
        [InlineData(10000)]
        [InlineData(30000)]
        [InlineData(32767)]
        public async Task Spreadsheet_AddRow_CellWithVeryLongStringValue(int length)
        {
            // Arrange
            var value = new string('a', length);
            using var stream = new MemoryStream();
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");
                var cell = new DataCell(value);

                // Act
                await spreadsheet.AddRowAsync(cell);
                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var actual = SpreadsheetDocument.Open(stream, true);
            var sheetPart = actual.WorkbookPart.WorksheetParts.Single();
            var actualCell = sheetPart.Worksheet.Descendants<Cell>().Single();
            Assert.Equal(CellValues.InlineString, actualCell.DataType.Value);
            Assert.Equal(value, actualCell.InnerText);
        }

        [Theory]
        [InlineData(1234)]
        [InlineData(0)]
        [InlineData(-1234)]
        [InlineData(int.MinValue)]
        [InlineData(int.MaxValue)]
        [InlineData(null)]
        public async Task Spreadsheet_AddRow_CellWithIntegerValue(int? value)
        {
            // Arrange
            using var stream = new MemoryStream();
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");
                var cell = new DataCell(value);

                // Act
                await spreadsheet.AddRowAsync(cell);
                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var actual = SpreadsheetDocument.Open(stream, true);
            var sheetPart = actual.WorkbookPart.WorksheetParts.Single();
            var actualCell = sheetPart.Worksheet.Descendants<Cell>().Single();
            Assert.Equal(CellValues.Number, actualCell.GetDataType());
            Assert.Equal(value?.ToString() ?? string.Empty, actualCell.InnerText);
        }

        [Theory]
        [InlineData(1234, "1234")]
        [InlineData(0, "0")]
        [InlineData(-1234, "-1234")]
        [InlineData(314748364700000, "314748364700000")]
        [InlineData(long.MinValue, "-9.22337203685478E+18")]
        [InlineData(long.MaxValue, "9.22337203685478E+18")]
        [InlineData(null, "")]
        public async Task Spreadsheet_AddRow_CellWithLongValue(long? initialValue, string expectedValue)
        {
            // Arrange
            using var stream = new MemoryStream();
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");
                var cell = new DataCell(initialValue);

                // Act
                await spreadsheet.AddRowAsync(cell);
                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var actual = SpreadsheetDocument.Open(stream, true);
            var sheetPart = actual.WorkbookPart.WorksheetParts.Single();
            var actualCell = sheetPart.Worksheet.Descendants<Cell>().Single();
            Assert.Equal(CellValues.Number, actualCell.GetDataType());
            Assert.Equal(expectedValue, actualCell.InnerText);
        }

        [Theory]
        [InlineData(1234f, "1234")]
        [InlineData(0.1f, "0.1")]
        [InlineData(0.0f, "0")]
        [InlineData(-0.1f, "-0.1")]
        [InlineData(0.1111111f, "0.1111111")]
        [InlineData(11.11111f, "11.11111")]
        [InlineData(2.222222E+38f, "2.222222E+38")]
        [InlineData(-0.3333333f, "-0.3333333")]
        [InlineData(null, "")]
        public async Task Spreadsheet_AddRow_CellWithFloatValue(float? initialValue, string expectedValue)
        {
            // Arrange
            using var stream = new MemoryStream();
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");
                var cell = new DataCell(initialValue);

                // Act
                await spreadsheet.AddRowAsync(cell);
                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var actual = SpreadsheetDocument.Open(stream, true);
            var sheetPart = actual.WorkbookPart.WorksheetParts.Single();
            var actualCell = sheetPart.Worksheet.Descendants<Cell>().Single();
            Assert.Equal(CellValues.Number, actualCell.GetDataType());
            Assert.Equal(expectedValue, actualCell.InnerText);
        }

        [Theory]
        [InlineData(1234, "1234")]
        [InlineData(0.1, "0.1")]
        [InlineData(0.0, "0")]
        [InlineData(-0.1, "-0.1")]
        [InlineData(0.1111111111111, "0.1111111111111")]
        [InlineData(11.1111111111111, "11.1111111111111")]
        [InlineData(11.11111111111111111111, "11.1111111111111")]
        [InlineData(2.2222222222E+50, "2.2222222222E+50")]
        [InlineData(-0.3333333, "-0.3333333")]
        [InlineData(null, "")]
        public async Task Spreadsheet_AddRow_CellWithDoubleValue(double? initialValue, string expectedValue)
        {
            // Arrange
            using var stream = new MemoryStream();
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");
                var cell = new DataCell(initialValue);

                // Act
                await spreadsheet.AddRowAsync(cell);
                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var actual = SpreadsheetDocument.Open(stream, true);
            var sheetPart = actual.WorkbookPart.WorksheetParts.Single();
            var actualCell = sheetPart.Worksheet.Descendants<Cell>().Single();
            Assert.Equal(CellValues.Number, actualCell.GetDataType());
            Assert.Equal(expectedValue, actualCell.InnerText);
        }

        [Theory]
        [InlineData("1234", "1234")]
        [InlineData("0.1", "0.1")]
        [InlineData("0.0", "0")]
        [InlineData("-0.1", "-0.1")]
        [InlineData("0.1111111111111", "0.1111111111111")]
        [InlineData("11.1111111111111", "11.1111111111111")]
        [InlineData("11.11111111111111111111", "11.1111111111111")]
        [InlineData("-0.3333333", "-0.3333333")]
        [InlineData("0.123456789012345678901234567", "0.123456789012346")]
        [InlineData(null, "")]
        public async Task Spreadsheet_AddRow_CellWithDecimalValue(string? initialValue, string expectedValue)
        {
            // Arrange
            var decimalValue = initialValue != null ? decimal.Parse(initialValue, CultureInfo.InvariantCulture) : null as decimal?;
            using var stream = new MemoryStream();
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");
                var cell = new DataCell(decimalValue);

                // Act
                await spreadsheet.AddRowAsync(cell);
                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var actual = SpreadsheetDocument.Open(stream, true);
            var sheetPart = actual.WorkbookPart.WorksheetParts.Single();
            var actualCell = sheetPart.Worksheet.Descendants<Cell>().Single();
            Assert.Equal(CellValues.Number, actualCell.GetDataType());
            Assert.Equal(expectedValue, actualCell.InnerText);
        }

        [Theory]
        [InlineData(true, "1")]
        [InlineData(false, "0")]
        [InlineData(null, "")]
        public async Task Spreadsheet_AddRow_CellWithBooleanValue(bool? initialValue, string expectedValue)
        {
            // Arrange
            using var stream = new MemoryStream();
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");
                var cell = new DataCell(initialValue);

                // Act
                await spreadsheet.AddRowAsync(cell);
                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var actual = SpreadsheetDocument.Open(stream, true);
            var sheetPart = actual.WorkbookPart.WorksheetParts.Single();
            var actualCell = sheetPart.Worksheet.Descendants<Cell>().Single();
            Assert.Equal(CellValues.Boolean, actualCell.GetDataType());
            Assert.Equal(expectedValue, actualCell.InnerText);
        }

        [Theory]
        [InlineData(1000)]
        public async Task Spreadsheet_AddRow_MultipleColumns(int columns)
        {
            // Arrange
            var values = Enumerable.Range(1, columns).Select(x => x.ToString()).ToList();
            using var stream = new MemoryStream();
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");
                var cells = values.Select(x => new DataCell(x)).ToList();

                // Act
                await spreadsheet.AddRowAsync(cells);
                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.Single();
            var actualValues = worksheet.Cells(true).Select(x => x.Value).ToList();
            Assert.Equal(values, actualValues);
        }

        [Theory]
        [InlineData(1000)]
        public async Task Spreadsheet_AddRow_MultipleRows(int rows)
        {
            // Arrange
            var values = Enumerable.Range(1, rows).Select(x => x.ToString()).ToList();
            using var stream = new MemoryStream();
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                // Act
                foreach (var value in values)
                {
                    await spreadsheet.AddRowAsync(new DataCell(value));
                }

                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.Single();
            var actualValues = worksheet.Cells(true).Select(x => x.Value).ToList();
            Assert.Equal(values, actualValues);
        }
    }
}
