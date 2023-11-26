using SpreadCheetah.Images;
using Xunit;

namespace SpreadCheetah.Test.Tests;

public class ImageCanvasTests
{
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
