using SpreadCheetah.Styling;

namespace SpreadCheetah.Test.Tests;

public class FontTests
{
    [Fact]
    public void Font_Name_TooLong()
    {
        // Arrange
        const string fontName = "FontNameThatIsExactly32CharsLong";
        var font = new Font();

        // Act & Assert
        Assert.Throws<ArgumentException>(() => font.Name = fontName);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(0.9)]
    [InlineData(409.1)]
    [InlineData(410)]
    public void Font_Size_OutOfRange(double size)
    {
        // Arrange
        var font = new Font();

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => font.Size = size);
    }

    [Fact]
    public void Font_Underline_InvalidValue()
    {
        // Arrange
        const Underline underline = (Underline)5;
        var font = new Font();

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => font.Underline = underline);
    }
}
