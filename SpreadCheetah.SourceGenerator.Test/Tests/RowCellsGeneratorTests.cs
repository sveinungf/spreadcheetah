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
        [Theory]
        [InlineData(ObjectType.Class)]
        [InlineData(ObjectType.Record)]
        [InlineData(ObjectType.Struct)]
        [InlineData(ObjectType.ReadOnlyStruct)]
        public async Task Spreadsheet_AddAsRow_ObjectWithProperties(ObjectType type)
        {
            // Arrange
            const string firstName = "Ola";
            const string lastName = "Nordmann";
            const int age = 30;

            using var stream = new MemoryStream();
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                // Act
                if (type == ObjectType.Class)
                    await spreadsheet.AddAsRowAsync(new ClassWithProperties(firstName, lastName, age));
                else if (type == ObjectType.Record)
                    await spreadsheet.AddAsRowAsync(new RecordWithProperties(firstName, lastName, age));
                else if (type == ObjectType.Struct)
                    await spreadsheet.AddAsRowAsync(new StructWithProperties(firstName, lastName, age));
                else if (type == ObjectType.ReadOnlyStruct)
                    await spreadsheet.AddAsRowAsync(new ReadOnlyStructWithProperties(firstName, lastName, age));

                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var actual = SpreadsheetDocument.Open(stream, false);
            var sheetPart = actual.WorkbookPart.WorksheetParts.Single();
            var cells = sheetPart.Worksheet.Descendants<Cell>().ToList();
            Assert.Equal(firstName, cells[0].InnerText);
            Assert.Equal(lastName, cells[1].InnerText);
            Assert.Equal(age.ToString(), cells[2].InnerText);
            Assert.Equal(3, cells.Count);
        }

        [Theory]
        [InlineData(ObjectType.Class)]
        [InlineData(ObjectType.Record)]
        [InlineData(ObjectType.Struct)]
        [InlineData(ObjectType.ReadOnlyStruct)]
        public async Task Spreadsheet_AddAsRow_ObjectWithNoProperties(ObjectType type)
        {
            // Arrange
            using var stream = new MemoryStream();
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                // Act
                if (type == ObjectType.Class)
                    await spreadsheet.AddAsRowAsync(new ClassWithNoProperties());
                else if (type == ObjectType.Record)
                    await spreadsheet.AddAsRowAsync(new RecordWithNoProperties());
                else if (type == ObjectType.Struct)
                    await spreadsheet.AddAsRowAsync(new StructWithNoProperties());
                else if (type == ObjectType.ReadOnlyStruct)
                    await spreadsheet.AddAsRowAsync(new ReadOnlyStructWithNoProperties());

                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var actual = SpreadsheetDocument.Open(stream, false);
            var sheetPart = actual.WorkbookPart.WorksheetParts.Single();
            Assert.Empty(sheetPart.Worksheet.Descendants<Cell>());
        }

        [Fact]
        public async Task Spreadsheet_AddAsRow_ObjectWithCustomType()
        {
            // Arrange
            const string value = "value";
            var customType = new CustomType("The name");
            using var stream = new MemoryStream();
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                // Act
                await spreadsheet.AddAsRowAsync(new RecordWithCustomType(customType, value));

                await spreadsheet.FinishAsync();
            }

            // Assert
            stream.Position = 0;
            using var actual = SpreadsheetDocument.Open(stream, false);
            var sheetPart = actual.WorkbookPart.WorksheetParts.Single();
            var cells = sheetPart.Worksheet.Descendants<Cell>().ToList();
            var actualCell = Assert.Single(cells);
            Assert.Equal(value, actualCell.InnerText);
        }
    }
}
