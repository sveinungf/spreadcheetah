using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadCheetah.SourceGenerator.Test.Models;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests
{
    public class RowCellsGeneratorTests
    {
        [Fact]
        public async Task Spreadsheet_AddAsRow_ClassWithProperties()
        {
            // Arrange
            var obj = new ClassWithProperties("Ola", "Nordmann", 30);
            using var stream = new MemoryStream();
            using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                // Act
                await spreadsheet.AddAsRowAsync(obj);
                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var actual = SpreadsheetDocument.Open(stream, false);
            var sheetPart = actual.WorkbookPart.WorksheetParts.Single();
            var cells = sheetPart.Worksheet.Descendants<Cell>().ToList();
            Assert.Equal(obj.FirstName, cells[0].InnerText);
            Assert.Equal(obj.LastName, cells[1].InnerText);
            Assert.Equal(obj.Age.ToString(), cells[2].InnerText);
            Assert.Equal(3, cells.Count);
        }
    }
}
