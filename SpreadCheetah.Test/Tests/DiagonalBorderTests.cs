using SpreadCheetah.Styling;

namespace SpreadCheetah.Test.Tests;

public class DiagonalBorderTests
{
    [Fact]
    public void DiagonalBorder_BorderStyle_InvalidValue()
    {
        // Arrange
        const BorderStyle borderStyle = (BorderStyle)14;
        var diagonalBorder = new DiagonalBorder();

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => diagonalBorder.BorderStyle = borderStyle);
    }

    [Fact]
    public void DiagonalBorder_Type_InvalidValue()
    {
        // Arrange
        const DiagonalBorderType type = (DiagonalBorderType)4;
        var diagonalBorder = new DiagonalBorder();

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => diagonalBorder.Type = type);
    }
}
