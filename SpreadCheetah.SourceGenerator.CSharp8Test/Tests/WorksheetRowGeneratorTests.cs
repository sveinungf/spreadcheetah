using SpreadCheetah.SourceGenerator.CSharp8Test.Models;
using SpreadCheetah.TestHelpers.Assertions;
using System;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace SpreadCheetah.SourceGenerator.CSharp8Test.Tests
{
    public class WorksheetRowGeneratorTests
    {
        [Fact]
        public async Task Spreadsheet_AddAsRow_ClassWithMultipleProperties()
        {
            // Arrange
            var obj = new ClassWithMultipleProperties(
                id: Guid.NewGuid().ToString(),
                firstName: "Ola",
                lastName: "Nordmann",
                age: 30);

            var ctx = ClassWithMultiplePropertiesContext.Default.ClassWithMultipleProperties;
            using var stream = new MemoryStream();
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet", ctx);

                // Act
                await spreadsheet.AddAsRowAsync(obj, ctx);

                await spreadsheet.FinishAsync();
            }

            // Assert
            using var sheet = SpreadsheetAssert.SingleSheet(stream);
            Assert.Equal(obj.FirstName, sheet["A1"].StringValue);
            Assert.Equal(obj.Id, sheet["B1"].StringValue);
            Assert.Equal(obj.LastName, sheet["C1"].StringValue);
            Assert.Equal(obj.Age, sheet["D1"].IntValue);
            Assert.Equal(4, sheet.CellCount);
            Assert.Equal(5, sheet.Column("D").Width, 4);
        }

        [Fact]
        public async Task Spreadsheet_AddAsRow_ClassWithNoProperties()
        {
            // Arrange
            var obj = new ClassWithNoProperties();

            using var stream = new MemoryStream();
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet");

                // Act
                await spreadsheet.AddAsRowAsync(obj, ClassWithNoPropertiesContext.Default.ClassWithNoProperties);

                await spreadsheet.FinishAsync();
            }

            // Assert
            using var sheet = SpreadsheetAssert.SingleSheet(stream);
            Assert.Equal(0, sheet.CellCount);
        }
    }
}
