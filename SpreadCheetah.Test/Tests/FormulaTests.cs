namespace SpreadCheetah.Test.Tests;

public static class FormulaTests
{
    [Fact]
    public static void Formula_Hyperlink_NullUri()
    {
        // Arrange
        Uri uri = null!;

        // Act & Assert
        Assert.Throws<ArgumentNullException>(() => Formula.Hyperlink(uri));
    }

    [Fact]
    public static void Formula_Hyperlink_RelativeUri()
    {
        // Arrange
        var uri = new Uri("/path/to/resource", UriKind.Relative);

        // Act & Assert
        Assert.Throws<ArgumentException>(() => Formula.Hyperlink(uri));
    }

    [Fact]
    public static void Formula_Hyperlink_NotWellFormedUri()
    {
        // Arrange
        var uri = new Uri("https://www.github.com/path???/file name", UriKind.Absolute);

        // Act & Assert
        Assert.Throws<ArgumentException>(() => Formula.Hyperlink(uri));
    }

    [Fact]
    public static void Formula_Hyperlink_TooLongUri()
    {
        // Arrange
        const string host = "https://www.github.com/";
        var path = new string('a', 256 - host.Length);
        var uri = new Uri(host + path, UriKind.Absolute);

        // Act & Assert
        Assert.Throws<ArgumentException>(() => Formula.Hyperlink(uri));
    }

    [Fact]
    public static void Formula_Hyperlink_NullFriendlyName()
    {
        // Arrange
        var uri = new Uri("https://www.github.com/");
        const string friendlyName = null!;

        // Act & Assert
        Assert.Throws<ArgumentNullException>(() => Formula.Hyperlink(uri, friendlyName));
    }

    [Fact]
    public static void Formula_Hyperlink_TooLongFriendlyName()
    {
        // Arrange
        var uri = new Uri("https://www.github.com/");
        var friendlyName = new string('a', 256);

        // Act & Assert
        Assert.Throws<ArgumentException>(() => Formula.Hyperlink(uri, friendlyName));
    }
}
