using SpreadCheetah.Styling;

namespace SpreadCheetah.Test.Tests;

public class AlignmentTests
{
    [Fact]
    public void Alignment_Indent_InvalidValue()
    {
        // Arrange
        var alignment = new Alignment();

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => alignment.Indent = -1);
    }
}
