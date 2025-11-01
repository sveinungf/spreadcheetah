using SpreadCheetah.SourceGenerator.Test.Models.InferColumnHeaders;
using SpreadCheetah.TestHelpers.Assertions;
using SpreadCheetah.TestHelpers.Extensions;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class InferColumnHeadersTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public async Task InferColumnHeaders_ClassWithMultipleProperties()
    {
        // Arrange
        var ctx = InferColumnHeadersContext.Default.ClassWithMultipleProperties;
        using var stream = new MemoryStream();
        await using var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await s.StartWorksheetAsync("Sheet", token: Token);

        string[] expectedValues =
        [
            "The ID",
            nameof(ClassWithMultipleProperties.Name),
            "The price"
        ];

        // Act
        await s.AddHeaderRowAsync(ctx, token: Token);
        await s.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(expectedValues, sheet.Row(1).Cells.StringValues());
    }

    [Fact]
    public async Task InferColumnHeaders_DerivedClassWithInferFromBaseClass()
    {
        // Arrange
        var ctx = InferColumnHeadersContext.Default.DerivedClassWithInferFromBaseClass;
        using var stream = new MemoryStream();
        await using var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await s.StartWorksheetAsync("Sheet", token: Token);

        string[] expectedValues =
        [
            "The make",
            "The model"
        ];

        // Act
        await s.AddHeaderRowAsync(ctx, token: Token);
        await s.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(expectedValues, sheet.Row(1).Cells.StringValues());
    }

    [Fact]
    public async Task InferColumnHeaders_DerivedClassWithInferFromBaseClassButNoInheritColumns()
    {
        // Arrange
        var ctx = InferColumnHeadersContext.Default.DerivedClassWithInferFromBaseClassButNoInheritColumns;
        using var stream = new MemoryStream();
        await using var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await s.StartWorksheetAsync("Sheet", token: Token);

        string[] expectedValues =
        [
            nameof(DerivedClassWithInferFromBaseClassButNoInheritColumns.Model)
        ];

        // Act
        await s.AddHeaderRowAsync(ctx, token: Token);
        await s.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(expectedValues, sheet.Row(1).Cells.StringValues());
    }

    [Fact]
    public async Task InferColumnHeaders_DerivedClassWithInfer()
    {
        // Arrange
        var ctx = InferColumnHeadersContext.Default.DerivedClassWithInfer;
        using var stream = new MemoryStream();
        await using var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await s.StartWorksheetAsync("Sheet", token: Token);

        string[] expectedValues =
        [
            nameof(BaseClassWithoutInfer.Make),
            "The model"
        ];

        // Act
        await s.AddHeaderRowAsync(ctx, token: Token);
        await s.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(expectedValues, sheet.Row(1).Cells.StringValues());
    }

    [Fact]
    public async Task InferColumnHeaders_DerivedClassWithInferAlsoFromBaseClass()
    {
        // Arrange
        var ctx = InferColumnHeadersContext.Default.DerivedClassWithInferAlsoFromBaseClass;
        using var stream = new MemoryStream();
        await using var s = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await s.StartWorksheetAsync("Sheet", token: Token);

        string[] expectedValues =
        [
            "The make",
            "The year"
        ];

        // Act
        await s.AddHeaderRowAsync(ctx, token: Token);
        await s.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(expectedValues, sheet.Row(1).Cells.StringValues());
    }
}
