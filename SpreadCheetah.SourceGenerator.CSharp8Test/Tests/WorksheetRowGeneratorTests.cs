using SpreadCheetah.SourceGenerator.CSharp8Test.Models;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Assertions;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace SpreadCheetah.SourceGenerator.CSharp8Test.Tests
{
    public class WorksheetRowGeneratorTests
    {
        private static CancellationToken Token => TestContext.Current.CancellationToken;

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
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
            {
                await spreadsheet.StartWorksheetAsync("Sheet", ctx, Token);
                spreadsheet.AddStyle(new Style { Font = { Bold = true } }, "Id style");
                spreadsheet.AddStyle(new Style { Font = { Italic = true } }, "Last name style");

                // Act
                await spreadsheet.AddAsRowAsync(obj, ctx, Token);

                await spreadsheet.FinishAsync(Token);
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
            Assert.Equal("#.0#", sheet["D1"].Style.NumberFormat.CustomFormat);
        }

        [Fact]
        public async Task Spreadsheet_AddAsRow_ClassWithNoProperties()
        {
            // Arrange
            var obj = new ClassWithNoProperties();

            using var stream = new MemoryStream();
            await using (var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token))
            {
                await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

                // Act
                await spreadsheet.AddAsRowAsync(obj, ClassWithNoPropertiesContext.Default.ClassWithNoProperties, Token);

                await spreadsheet.FinishAsync(Token);
            }

            // Assert
            using var sheet = SpreadsheetAssert.SingleSheet(stream);
            Assert.Equal(0, sheet.CellCount);
        }
    }
}
