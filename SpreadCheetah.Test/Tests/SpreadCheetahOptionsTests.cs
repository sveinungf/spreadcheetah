namespace SpreadCheetah.Test.Tests;

public class SpreadCheetahOptionsTests
{
    [Fact]
    public void SpreadCheetahOptions_BufferSize_Invalid()
    {
        // Arrange
        var options = new SpreadCheetahOptions();

        // Act & Assert
        Assert.ThrowsAny<ArgumentOutOfRangeException>(() => options.BufferSize = 500);
    }

#pragma warning disable CS0618 // Type or member is obsolete
    [Fact]
    public void SpreadCheetahOptions_DefaultDateTimeNumberFormat_Success()
    {
        // Arrange
        const string format = "[h] : mm : ss";
        var options = new SpreadCheetahOptions();

        // Act
        options.DefaultDateTimeNumberFormat = format;

        // Assert
        Assert.Equal(format, options.DefaultDateTimeNumberFormat);
    }
#pragma warning restore CS0618 // Type or member is obsolete
}
