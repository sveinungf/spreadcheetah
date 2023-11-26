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
        Assert.ThrowsAny<ArgumentOutOfRangeException>(() => ImageCanvas.Dimensions("A1".AsSpan(), width, height));
    }

    [Theory]
    [InlineData(-0.1)]
    [InlineData(0.0)]
    [InlineData(1001)]
    public void ImageCanvas_Scale_Invalid(decimal scale)
    {
        Assert.ThrowsAny<ArgumentOutOfRangeException>(() => ImageCanvas.Scaled("A1".AsSpan(), scale));
    }

    [Theory]
    [InlineData("A2")]
    [InlineData("B3")]
    [InlineData("B4")]
    [InlineData("C3")]
    public void ImageCanvas_FillCells_InvalidCellRange(string lowerRightReference)
    {
        // Act & Assert
        Assert.ThrowsAny<ArgumentOutOfRangeException>(() => ImageCanvas.FillCells("B3".AsSpan(), lowerRightReference.AsSpan()));
    }
}
