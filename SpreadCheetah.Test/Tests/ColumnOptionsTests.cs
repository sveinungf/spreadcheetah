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

    [Theory]
    [InlineData(-1)]
    [InlineData(8)]
    public void ColumnOptions_OutlineLevel_Invalid(int outlineLevel)
    {
        // Arrange
        var options = new ColumnOptions();

        // Act
        var exception = Record.Exception(() => options.OutlineLevel = outlineLevel);

        // Assert
        Assert.IsType<ArgumentOutOfRangeException>(exception);
    }
}
