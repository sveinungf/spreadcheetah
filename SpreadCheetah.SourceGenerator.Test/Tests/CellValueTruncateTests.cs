using SpreadCheetah.SourceGeneration;
using SpreadCheetah.SourceGenerator.Test.Helpers;
using SpreadCheetah.SourceGenerator.Test.Models.CellValueTruncation;
using SpreadCheetah.SourceGenerator.Test.Models.Contexts;
using SpreadCheetah.TestHelpers.Assertions;
using System.Reflection;
using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Tests;

public class CellValueTruncateTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Theory]
    [InlineData("Short value")]
    [InlineData("Exact length!!!")]
    [InlineData("Long value that will be truncated")]
#pragma warning disable S2479 // Whitespace and control characters in string literals should be explicit
    [InlineData("A couple 👨‍👩‍👧‍👦 with kids")]
#pragma warning restore S2479 // Whitespace and control characters in string literals should be explicit
    [InlineData("")]
    [InlineData(null)]
    public async Task CellValueTruncate_ClassWithAttribute(string? originalValue)
    {
        // Arrange
        const int truncateLength = 15;
        var expectedValue = originalValue is { Length: > truncateLength }
            ? originalValue[..truncateLength]
            : originalValue;

        var obj = new ClassWithTruncation { Value = originalValue };
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        // Act
        await spreadsheet.AddAsRowAsync(obj, TruncationContext.Default.ClassWithTruncation, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal(expectedValue, sheet["A1"].StringValue);
    }

    [Fact]
    public async Task CellValueTruncate_ClassWithSingleAccessProperty()
    {
        // Arrange
        var obj = new ClassWithSingleAccessProperty("The value");
        using var stream = new MemoryStream();
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(stream, cancellationToken: Token);
        await spreadsheet.StartWorksheetAsync("Sheet", token: Token);

        // Act
        await spreadsheet.AddAsRowAsync(obj, TruncationContext.Default.ClassWithSingleAccessProperty, Token);
        await spreadsheet.FinishAsync(Token);

        // Assert
        using var sheet = SpreadsheetAssert.SingleSheet(stream);
        Assert.Equal("T", sheet["A1"].StringValue);
    }

    [Fact]
    public void CellValueTruncate_ClassWithSingleAccessProperty_CanReadLength()
    {
        // Arrange
        var properties = typeof(ClassWithSingleAccessProperty).ToPropertyDictionary();
        var property = properties[nameof(ClassWithSingleAccessProperty.Value)];

        // Act
        var attribute = property.GetCustomAttribute<CellValueTruncateAttribute>();

        // Assert
        Assert.NotNull(attribute);
        Assert.Equal(1, attribute.Length);
    }
}
