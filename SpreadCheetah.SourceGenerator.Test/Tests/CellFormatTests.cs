using SpreadCheetah.SourceGenerator.Test.Models.CellFormat;
using SpreadCheetah.Styling;
using SpreadCheetah.TestHelpers.Assertions;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class CellFormatTests
{
    [Fact]
    public async Task CellFormat_ClassWithStandardFormat()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream);
        await spreadsheet.StartWorksheetAsync("Sheet");
        var obj = new ClassWithCellStandardFormat { Price = 199.90m };

        // Act
        await spreadsheet.AddAsRowAsync(obj, CellFormatContext.Default.ClassWithCellStandardFormat);
        await spreadsheet.FinishAsync();

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(obj.Price, sheet["A1"].DecimalValue);
        Assert.Equal(StandardNumberFormat.TwoDecimalPlaces, sheet["A1"].Style.NumberFormat.StandardFormat);
    }
}
