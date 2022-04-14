using DocumentFormat.OpenXml.Packaging;
using SpreadCheetah.SourceGenerator.CSharp8Test.Models;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;
using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;

namespace SpreadCheetah.SourceGenerator.CSharp8Test.Tests
{
    public class WorksheetRowGeneratorTests
    {
        [Fact]
        public async Task Spreadsheet_AddAsRow_ClassWithMultipleProperties()
        {
            // Arrange
            const string firstName = "Ola";
            const string lastName = "Nordmann";
            const int age = 30;
            var obj = new ClassWithMultipleProperties(firstName, lastName, age);

            using var stream = new MemoryStream();
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                // Act
                await spreadsheet.AddAsRowAsync(obj, ClassWithMultiplePropertiesContext.Default.ClassWithMultipleProperties);

                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var actual = SpreadsheetDocument.Open(stream, false);
            var sheetPart = actual.WorkbookPart.WorksheetParts.Single();
            var cells = sheetPart.Worksheet.Descendants<OpenXmlCell>().ToList();
            Assert.Equal(firstName, cells[0].InnerText);
            Assert.Equal(lastName, cells[1].InnerText);
            Assert.Equal(age.ToString(), cells[2].InnerText);
            Assert.Equal(3, cells.Count);
        }
    }
}
