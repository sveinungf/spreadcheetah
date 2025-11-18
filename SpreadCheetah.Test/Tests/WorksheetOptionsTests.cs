using SpreadCheetah.Worksheets;

namespace SpreadCheetah.Test.Tests;

public class WorksheetOptionsTests
{
    [Theory]
    [InlineData(0)]
    [InlineData(16385)]
    public void WorksheetOptions_FrozenColumns_Invalid(int value)
    {
        // Arrange
        var options = new WorksheetOptions();

        // Act
        var exception = Record.Exception(() => options.FrozenColumns = value);

        // Assert
        Assert.IsType<ArgumentOutOfRangeException>(exception);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1048577)]
    public void WorksheetOptions_FrozenRows_Invalid(int value)
    {
        // Arrange
        var options = new WorksheetOptions();

        // Act
        var exception = Record.Exception(() => options.FrozenRows = value);

        // Assert
        Assert.IsType<ArgumentOutOfRangeException>(exception);
    }
}
