using SpreadCheetah.SourceGenerator.Test.Models.Formulas;
using SpreadCheetah.Styling;
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
        const string formulaString = "SUM(C1:C10)";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var obj = new ClassWithFormula { MyFormula = new Formula(formulaString) };

        // Act
        await spreadsheet.AddAsRowAsync(obj, FormulaContext.Default.ClassWithFormula, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(formulaString, sheet["A1"].Formula);
    }

    [Theory, CombinatorialData]
    public async Task Spreadsheet_AddAsRow_ClassWithNullableFormula(bool isNull)
    {
        // Arrange
        const string formulaString = "SUM(C1:C10)";
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        var obj = new ClassWithNullableFormula { MyFormula = isNull ? null : new Formula(formulaString) };

        // Act
        await spreadsheet.AddAsRowAsync(obj, FormulaContext.Default.ClassWithNullableFormula, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(isNull ? "" : formulaString, sheet["A1"].Formula);
    }

    [Fact]
    public async Task Spreadsheet_AddAsRow_ClassWithStyledFormulas()
    {
        // Arrange
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);
        spreadsheet.AddStyle(new Style { Font = { Bold = true } }, "Bold");
        var obj = new ClassWithStyledFormulas
        {
            BoldFormula = new Formula("SUM(C1:C10)"),
            FormatFormula = new Formula("AVERAGE(C1:C10)")
        };

        // Act
        await spreadsheet.AddAsRowAsync(obj, FormulaContext.Default.ClassWithStyledFormulas, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.True(sheet["A1"].Style.Font.Bold);
        Assert.Equal("#.##", sheet["B1"].Style.NumberFormat.CustomFormat);
    }
}
