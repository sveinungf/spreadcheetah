using SpreadCheetah.SourceGenerator.CSharp8Test.Models;
using SpreadCheetah.Styling;
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
            const string id = "2cc56ae0-32b2-46c2-b6d2-27c6ce8568ec";
            var obj = new ClassWithMultipleProperties(
                id: id,
                firstName: "Ola",
                lastName: "Nordmann",
                age: 30);

            var ctx = ClassWithMultiplePropertiesContext.Default.ClassWithMultipleProperties;
            using var stream = new MemoryStream();
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream))
            {
                await spreadsheet.StartWorksheetAsync("Sheet", ctx);
                spreadsheet.AddStyle(new Style { Font = { Bold = true } }, "Id style");
                spreadsheet.AddStyle(new Style { Font = { Italic = true } }, "Last name style");

                // Act
                await spreadsheet.AddAsRowAsync(obj, ctx);

                await spreadsheet.FinishAsync();
            }

            // Assert
            using var sheet = SpreadsheetAssert.SingleSheet(stream);
            Assert.Equal(obj.FirstName.ToUpperInvariant(), sheet["A1"].StringValue);
            Assert.Equal(obj.Id.Replace("-", ""), sheet["B1"].StringValue);
            Assert.Equal(obj.LastName, sheet["C1"].StringValue);
            Assert.Equal(obj.Age, sheet["D1"].IntValue);
            Assert.Equal(4, sheet.CellCount);
            Assert.Equal(5, sheet.Column("D").Width, 4);

            Assert.True(sheet["B1"].Style.Font.Bold);
            Assert.True(sheet["C1"].Style.Font.Italic);
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
