using ClosedXML.Excel;
using SpreadCheetah.Test.Helpers;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;

namespace SpreadCheetah.Test.Tests
{
    public class SpreadsheetStyledRowTests
    {
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
        public async Task Spreadsheet_AddRow_NullStyleIdCellWithStringValue(string? value)
        {
            // Arrange
            using var stream = new MemoryStream();
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");
                var cell = new Cell(value);
                var styledCell = new StyledCell(cell, null);

                // Act
                await spreadsheet.AddRowAsync(styledCell);
                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.Single();
            var actualCell = worksheet.Cell(1, 1);
            Assert.Equal(value ?? string.Empty, actualCell.Value);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public async Task Spreadsheet_AddRow_BoldCellWithStringValue(bool bold)
        {
            // Arrange
            const string cellValue = "Bold test";
            using var stream = new MemoryStream();
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                var style = new Style();
                style.Font.Bold = bold;
                var styleId = spreadsheet.AddStyle(style);

                var cell = new Cell(cellValue);
                var styledCell = new StyledCell(cell, styleId);

                // Act
                await spreadsheet.AddRowAsync(styledCell);
                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.Single();
            var actualCell = worksheet.Cell(1, 1);
            Assert.Equal(cellValue, actualCell.Value);
            Assert.Equal(bold, actualCell.Style.Font.Bold);
        }
    }
}
