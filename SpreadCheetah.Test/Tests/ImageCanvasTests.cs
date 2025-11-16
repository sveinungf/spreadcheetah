using SpreadCheetah.Images;
using Xunit;

namespace SpreadCheetah.Test.Tests;

public class ImageCanvasTests
{
    [Theory]
    [InlineData(0, 100)]
    [InlineData(100, 0)]
    [InlineData(100, 100001)]
    [InlineData(100001, 100)]
    [InlineData(-100, 100)]
    [InlineData(100, -100)]
    public void ImageCanvas_Dimensions_Invalid(int width, int height)
    {
        Assert.ThrowsAny<ArgumentOutOfRangeException>(() => ImageCanvas.Dimensions("A1", width, height));
    }

    [Theory]
    [InlineData(-0.1)]
    [InlineData(0.0)]
    [InlineData(1001)]
    public void ImageCanvas_Scaled_Invalid(float scale)
    {
        Assert.ThrowsAny<ArgumentOutOfRangeException>(() => ImageCanvas.Scaled("A1", scale));
    }

    [Fact]
    public void ImageCanvas_FillCell_InvalidAnchor()
    {
        // Act
        var exception = Record.Exception(() => ImageCanvas.FillCell("A1", moveWithCells: false, resizeWithCells: true));

        // Assert
        var concreteException = Assert.IsType<ArgumentException>(exception);
        Assert.Contains("resize", concreteException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ImageCanvas_FillCells_InvalidAnchor()
    {
        // Act
        var exception = Record.Exception(() => ImageCanvas.FillCells("A1", "D4", moveWithCells: false, resizeWithCells: true));

        // Assert
        var concreteException = Assert.IsType<ArgumentException>(exception);
        Assert.Contains("resize", concreteException.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("A2")]
    [InlineData("B3")]
    [InlineData("B4")]
    [InlineData("C3")]
    public void ImageCanvas_FillCells_InvalidCellRange(string lowerRightReference)
    {
        // Act & Assert
        Assert.ThrowsAny<ArgumentOutOfRangeException>(() => ImageCanvas.FillCells("B3", lowerRightReference));
    }
}
