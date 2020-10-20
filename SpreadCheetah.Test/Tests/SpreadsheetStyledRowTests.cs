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

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public async Task Spreadsheet_AddRow_SameBoldStyleCells(bool bold)
        {
            // Arrange
            const string firstCellValue = "First";
            const string secondCellValue = "Second";
            using var stream = new MemoryStream();
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                var style = new Style();
                style.Font.Bold = bold;
                var styleId = spreadsheet.AddStyle(style);

                var firstCell = new StyledCell(new Cell(firstCellValue), styleId);
                var secondCell = new StyledCell(new Cell(secondCellValue), styleId);

                // Act
                await spreadsheet.AddRowAsync(new[] { firstCell, secondCell });
                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.Single();
            var actualFirstCell = worksheet.Cell(1, 1);
            var actualSecondCell = worksheet.Cell(1, 2);
            Assert.Equal(firstCellValue, actualFirstCell.Value);
            Assert.Equal(secondCellValue, actualSecondCell.Value);
            Assert.Equal(bold, actualFirstCell.Style.Font.Bold);
            Assert.Equal(bold, actualSecondCell.Style.Font.Bold);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public async Task Spreadsheet_AddRow_MixedBoldStyleCells(bool firstCellBold)
        {
            // Arrange
            const string firstCellValue = "First";
            const string secondCellValue = "Second";
            using var stream = new MemoryStream();
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                var style = new Style();
                style.Font.Bold = true;
                var styleId = spreadsheet.AddStyle(style);

                var firstCell = new StyledCell(new Cell(firstCellValue), firstCellBold ? styleId : null);
                var secondCell = new StyledCell(new Cell(secondCellValue), firstCellBold ? null : styleId);

                // Act
                await spreadsheet.AddRowAsync(new[] { firstCell, secondCell });
                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.Single();
            var actualFirstCell = worksheet.Cell(1, 1);
            var actualSecondCell = worksheet.Cell(1, 2);
            Assert.Equal(firstCellValue, actualFirstCell.Value);
            Assert.Equal(secondCellValue, actualSecondCell.Value);
            Assert.Equal(firstCellBold, actualFirstCell.Style.Font.Bold);
            Assert.Equal(!firstCellBold, actualSecondCell.Style.Font.Bold);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public async Task Spreadsheet_AddRow_MixedCellTypeRows(bool firstRowStyled)
        {
            // Arrange
            const string firstCellValue = "First";
            const string secondCellValue = "Second";
            using var stream = new MemoryStream();
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                var style = new Style();
                style.Font.Bold = true;
                var styleId = spreadsheet.AddStyle(style);

                var firstCell = new Cell(firstCellValue);
                var secondCell = new Cell(secondCellValue);

                // Act
                if (firstRowStyled)
                {
                    await spreadsheet.AddRowAsync(new StyledCell(firstCell, styleId));
                    await spreadsheet.AddRowAsync(secondCell);
                }
                else
                {
                    await spreadsheet.AddRowAsync(firstCell);
                    await spreadsheet.AddRowAsync(new StyledCell(secondCell, styleId));
                }

                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.Single();
            var actualFirstCell = worksheet.Cell(1, 1);
            var actualSecondCell = worksheet.Cell(2, 1);
            Assert.Equal(firstCellValue, actualFirstCell.Value);
            Assert.Equal(secondCellValue, actualSecondCell.Value);
            Assert.Equal(firstRowStyled, actualFirstCell.Style.Font.Bold);
            Assert.Equal(!firstRowStyled, actualSecondCell.Style.Font.Bold);
        }
    }
}
