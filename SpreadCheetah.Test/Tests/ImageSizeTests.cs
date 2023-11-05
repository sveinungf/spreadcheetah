using SpreadCheetah.Images;
using Xunit;

namespace SpreadCheetah.Test.Tests;

public class ImageSizeTests
{
    [Theory]
    [InlineData(0, 100)]
    [InlineData(100, 0)]
    [InlineData(100, 100001)]
    [InlineData(100001, 100)]
    [InlineData(-100, 100)]
    [InlineData(100, -100)]
    public void ImageSize_Dimensions_Invalid(int width, int height)
    {
        Assert.ThrowsAny<ArgumentOutOfRangeException>(() => ImageSize.Dimensions(width, height));
    }
}
