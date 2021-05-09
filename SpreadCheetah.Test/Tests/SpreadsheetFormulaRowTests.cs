using ClosedXML.Excel;
using SpreadCheetah.Test.Helpers;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;

namespace SpreadCheetah.Test.Tests
{
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
            stream.Position = 0;
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.Single();
            var actualCell = worksheet.Cell(1, 1);
            Assert.Equal(formulaText, actualCell.FormulaA1);
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
            stream.Position = 0;
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.Single();
            var actualCell = worksheet.Cell(1, 1);
            Assert.Equal(formulaText, actualCell.FormulaA1);
        }

        [Theory]
        [InlineData(100)]
        [InlineData(511)]
        [InlineData(512)]
        [InlineData(513)]
        [InlineData(4100)]
        [InlineData(8192)]
        public async Task Spreadsheet_AddRow_CellWithVeryLongFormulaAndCachedValue(int length)
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
            stream.Position = 0;
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.Single();
            var actualCell = worksheet.Cell(1, 1);
            Assert.Equal(formulaText, actualCell.FormulaA1);
            Assert.Equal(cachedValue, actualCell.Value);
        }
    }
}
