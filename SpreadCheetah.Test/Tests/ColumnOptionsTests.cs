using SpreadCheetah.Worksheets;

namespace SpreadCheetah.Test.Tests;

public class ColumnOptionsTests
{
    [Theory]
    [InlineData(-1)]
    [InlineData(256)]
    public void ColumnOptions_Width_Invalid(double width)
    {
        // Arrange
        var options = new ColumnOptions();

        // Act
        var exception = Record.Exception(() => options.Width = width);

        // Assert
        Assert.IsType<ArgumentOutOfRangeException>(exception);
    }
}
