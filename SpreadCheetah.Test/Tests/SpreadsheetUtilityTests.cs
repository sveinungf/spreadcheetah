using System.Text;
using Xunit;

namespace SpreadCheetah.Test.Tests;

public static class SpreadsheetUtilityTests
{
    [Theory]
    [InlineData(1, "A")]
    [InlineData(3, "C")]
    [InlineData(24, "X")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(29, "AC")]
    [InlineData(50, "AX")]
    [InlineData(52, "AZ")]
    [InlineData(53, "BA")]
    [InlineData(55, "BC")]
    [InlineData(700, "ZX")]
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    [InlineData(705, "AAC")]
    [InlineData(728, "AAZ")]
    [InlineData(729, "ABA")]
    [InlineData(731, "ABC")]
    [InlineData(1378, "AZZ")]
    [InlineData(1379, "BAA")]
    [InlineData(1381, "BAC")]
    [InlineData(5000, "GJH")]
    [InlineData(10000, "NTP")]
    [InlineData(16384, "XFD")]
    public static void SpreadsheetUtility_GetColumnName_Success(int number, string expectedName)
    {
        // Act
        var actualName = SpreadsheetUtility.GetColumnName(number);

        // Assert
        Assert.Equal(expectedName, actualName);
    }

    [Theory]
    [InlineData(1, "A")]
    [InlineData(3, "C")]
    [InlineData(24, "X")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(29, "AC")]
    [InlineData(50, "AX")]
    [InlineData(52, "AZ")]
    [InlineData(53, "BA")]
    [InlineData(55, "BC")]
    [InlineData(700, "ZX")]
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    [InlineData(705, "AAC")]
    [InlineData(728, "AAZ")]
    [InlineData(729, "ABA")]
    [InlineData(731, "ABC")]
    [InlineData(1378, "AZZ")]
    [InlineData(1379, "BAA")]
    [InlineData(1381, "BAC")]
    [InlineData(5000, "GJH")]
    [InlineData(10000, "NTP")]
    [InlineData(16384, "XFD")]
    public static void SpreadsheetUtility_TryGetColumnNameUtf8_Success(int number, string expectedName)
    {
        // Arrange
        var bytes = new byte[10];

        // Act
        var result = SpreadsheetUtility.TryGetColumnNameUtf8(number, bytes, out var bytesWritten);

        // Assert
        Assert.True(result);
        var actualName = Encoding.UTF8.GetString(bytes, 0, bytesWritten);
        Assert.Equal(expectedName, actualName);
        Assert.All(bytes.Skip(bytesWritten), x => Assert.Equal((byte)0, x));
    }

    [Theory]
    [InlineData(1, "A")]
    [InlineData(3, "C")]
    [InlineData(24, "X")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(29, "AC")]
    [InlineData(50, "AX")]
    [InlineData(52, "AZ")]
    [InlineData(53, "BA")]
    [InlineData(55, "BC")]
    [InlineData(700, "ZX")]
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    [InlineData(705, "AAC")]
    [InlineData(728, "AAZ")]
    [InlineData(729, "ABA")]
    [InlineData(731, "ABC")]
    [InlineData(1378, "AZZ")]
    [InlineData(1379, "BAA")]
    [InlineData(1381, "BAC")]
    [InlineData(5000, "GJH")]
    [InlineData(10000, "NTP")]
    [InlineData(16384, "XFD")]
    public static void SpreadsheetUtility_TryParseColumnName_Success(int number, string name)
    {
        // Act
        var result = SpreadsheetUtility.TryParseColumnName(name.AsSpan(), out var actualNumber);

        // Assert
        Assert.True(result);
        Assert.Equal(number, actualNumber);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("@")] // Character before 'A'
    [InlineData("[")] // Character after 'Z'
    [InlineData("a")]
    [InlineData(" A")]
    [InlineData("A ")]
    [InlineData("Aa")]
    [InlineData("AAa")]
    [InlineData("AAAA")]
    [InlineData("XFE")] // XFD is the last allowed column name
    public static void SpreadsheetUtility_TryParseColumnName_Invalid(string name)
    {
        // Act
        var result = SpreadsheetUtility.TryParseColumnName(name.AsSpan(), out var columnNumber);

        // Assert
        Assert.False(result);
        Assert.Equal(0, columnNumber);
    }
}
