using SpreadCheetah.Styling;

namespace SpreadCheetah.Test.Tests;

public class EdgeBorderTests
{
    [Fact]
    public void EdgeBorder_BorderStyle_InvalidValue()
    {
        // Arrange
        const BorderStyle borderStyle = (BorderStyle)14;
        var edgeBorder = new EdgeBorder();

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => edgeBorder.BorderStyle = borderStyle);
    }
}
