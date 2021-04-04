using ClosedXML.Excel;
using SpreadCheetah.Styling;
using SpreadCheetah.Test.Helpers;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;

namespace SpreadCheetah.Test.Tests
{
    public class SpreadsheetStyledRowTests
    {
        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public async Task Spreadsheet_AddRow_BoldCellWithStringValue(bool bold)
        {
            // Arrange
            const string cellValue = "Bold test";
            using var stream = new MemoryStream();
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                var style = new Style();
                style.Font.Bold = bold;
                var styleId = spreadsheet.AddStyle(style);
                var styledCell = new StyledCell(cellValue, styleId);

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
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                var style = new Style();
                style.Font.Bold = bold;
                var styleId = spreadsheet.AddStyle(style);

                var firstCell = new StyledCell(firstCellValue, styleId);
                var secondCell = new StyledCell(secondCellValue, styleId);

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
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                var style = new Style();
                style.Font.Bold = true;
                var styleId = spreadsheet.AddStyle(style);

                var firstCell = new StyledCell(firstCellValue, firstCellBold ? styleId : null);
                var secondCell = new StyledCell(secondCellValue, firstCellBold ? null : styleId);

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
        public async Task Spreadsheet_AddRow_ItalicCellWithStringValue(bool italic)
        {
            // Arrange
            const string cellValue = "Italic test";
            using var stream = new MemoryStream();
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                var style = new Style();
                style.Font.Italic = italic;
                var styleId = spreadsheet.AddStyle(style);
                var styledCell = new StyledCell(cellValue, styleId);

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
            Assert.Equal(italic, actualCell.Style.Font.Italic);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public async Task Spreadsheet_AddRow_StrikethroughCellWithStringValue(bool strikethrough)
        {
            // Arrange
            const string cellValue = "Italic test";
            using var stream = new MemoryStream();
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                var style = new Style();
                style.Font.Strikethrough = strikethrough;
                var styleId = spreadsheet.AddStyle(style);
                var styledCell = new StyledCell(cellValue, styleId);

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
            Assert.Equal(strikethrough, actualCell.Style.Font.Strikethrough);
        }

        [Theory]
        [InlineData(8)]
        [InlineData(11)]
        [InlineData(11.5)]
        [InlineData(72)]
        public async Task Spreadsheet_AddRow_FontSizeCellWithStringValue(double size)
        {
            // Arrange
            const string cellValue = "Font size test";
            using var stream = new MemoryStream();
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                var style = new Style();
                style.Font.Size = size;
                var styleId = spreadsheet.AddStyle(style);
                var styledCell = new StyledCell(cellValue, styleId);

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
            Assert.Equal(size, actualCell.Style.Font.FontSize);
        }

        [Fact]
        public async Task Spreadsheet_AddRow_FontColorCellWithStringValue()
        {
            // Arrange
            const string cellValue = "Color test";
            var color = Color.Blue;
            using var stream = new MemoryStream();
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                var style = new Style();
                style.Font.Color = color;
                var styleId = spreadsheet.AddStyle(style);
                var styledCell = new StyledCell(cellValue, styleId);

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
            Assert.Equal(color, actualCell.Style.Font.FontColor.Color);
        }

        [Fact]
        public async Task Spreadsheet_AddRow_FillColorCellWithStringValue()
        {
            // Arrange
            const string cellValue = "Color test";
            var color = Color.Brown;
            using var stream = new MemoryStream();
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                var style = new Style();
                style.Fill.Color = color;
                var styleId = spreadsheet.AddStyle(style);
                var styledCell = new StyledCell(cellValue, styleId);

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
            Assert.Equal(color, actualCell.Style.Fill.BackgroundColor.Color);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public async Task Spreadsheet_AddRow_MultiFormatCellWithStringValue(bool formatting)
        {
            // Arrange
            const string cellValue = "Formatting test";
            var fillColor = formatting ? Color.Green as Color? : null;
            var fontColor = formatting ? Color.Red as Color? : null;
            using var stream = new MemoryStream();
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                var style = new Style();
                style.Fill.Color = fillColor;
                style.Font.Bold = formatting;
                style.Font.Color = fontColor;
                style.Font.Italic = formatting;
                style.Font.Strikethrough = formatting;
                var styleId = spreadsheet.AddStyle(style);
                var styledCell = new StyledCell(cellValue, styleId);

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
            Assert.Equal(formatting, actualCell.Style.Fill.BackgroundColor.Color == fillColor);
            Assert.Equal(formatting, actualCell.Style.Font.Bold);
            Assert.Equal(formatting, actualCell.Style.Font.FontColor.Color == fontColor);
            Assert.Equal(formatting, actualCell.Style.Font.Italic);
            Assert.Equal(formatting, actualCell.Style.Font.Strikethrough);
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
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                var style = new Style();
                style.Font.Bold = true;
                var styleId = spreadsheet.AddStyle(style);

                // Act
                if (firstRowStyled)
                {
                    await spreadsheet.AddRowAsync(new StyledCell(firstCellValue, styleId));
                    await spreadsheet.AddRowAsync(new DataCell(secondCellValue));
                }
                else
                {
                    await spreadsheet.AddRowAsync(new DataCell(firstCellValue));
                    await spreadsheet.AddRowAsync(new StyledCell(secondCellValue, styleId));
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
