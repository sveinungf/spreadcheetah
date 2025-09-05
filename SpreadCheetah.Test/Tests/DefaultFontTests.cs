using SpreadCheetah.Styling;

namespace SpreadCheetah.Test.Tests;

public class DefaultFontTests
{
    [Theory]
    [InlineData(2)]
    [InlineData(32)]
    public void DefaultFont_Name_InvalidLength(int length)
    {
        // Arrange
        var fontName = new string('a', length);
        var font = new DefaultFont();

        // Act & Assert
        Assert.Throws<ArgumentException>(() => font.Name = fontName);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(0.9)]
    [InlineData(409.1)]
    [InlineData(410)]
    public void DefaultFont_Size_OutOfRange(double size)
    {
        // Arrange
        var font = new DefaultFont();

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => font.Size = size);
    }
}
