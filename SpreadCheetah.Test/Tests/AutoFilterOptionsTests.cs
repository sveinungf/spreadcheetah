using SpreadCheetah.Worksheets;
using Xunit;

namespace SpreadCheetah.Test.Tests;

public static class AutoFilterOptionsTests
{
    [Theory]
    [InlineData("A1:A2")]
    [InlineData("A1:F1")]
    [InlineData("A1:XFD10000")]
    public static void AutoFilterOptions_Create_ValidRange(string cellRange)
    {
        // Act
        var exception = Record.Exception(() => new AutoFilterOptions(cellRange));

        // Assert
        Assert.Null(exception);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("A")]
    [InlineData("A1")]
    [InlineData("A1:A")]
    [InlineData("A1:AAAA2")]
    [InlineData("$A$1:$A$2")]
    public static void AutoFilterOptions_Create_InvalidRange(string cellRange)
    {
        // Act & Assert
        Assert.ThrowsAny<ArgumentException>(() => new AutoFilterOptions(cellRange));
    }
}
