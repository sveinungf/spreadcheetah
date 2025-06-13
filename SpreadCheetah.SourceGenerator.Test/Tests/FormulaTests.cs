using SpreadCheetah.SourceGenerator.Test.Models.Formulas;
using SpreadCheetah.TestHelpers.Assertions;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class FormulaTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public async Task Spreadsheet_AddAsRow_ClassWithFormula()
    {
        // Arrange
        const string formulaString = "SUM(A1:A10)";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var obj = new ClassWithFormula { MyFormula = new Formula(formulaString) };

        // Act
        await spreadsheet.AddAsRowAsync(obj, FormulaContext.Default.ClassWithFormula, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal("SUM(A1:A10)", sheet["A1"].Formula);
    }
}
