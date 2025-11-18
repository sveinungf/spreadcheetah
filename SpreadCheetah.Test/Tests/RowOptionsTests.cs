using SpreadCheetah.Worksheets;

namespace SpreadCheetah.Test.Tests;

public class RowOptionsTests
{
    [Theory]
    [InlineData(0)]
    [InlineData(410)]
    public void RowOptions_Height_Invalid(double value)
    {
        // Arrange
        var options = new RowOptions();

        // Act
        var exception = Record.Exception(() => options.Height = value);

        // Assert
        Assert.IsType<ArgumentOutOfRangeException>(exception);
    }
}
