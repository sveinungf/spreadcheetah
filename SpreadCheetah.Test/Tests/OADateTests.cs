#if DEBUG
using SpreadCheetah.Helpers;
using SpreadCheetah.Test.Helpers;
using System.Globalization;
using System.Text;

namespace SpreadCheetah.Test.Tests;

public class OADateTests
{
    public static TheoryData<DateTime> DateTimes => new()
    {
        { DateTime.MinValue },
        { new DateTime(1, 1, 1, 23, 59, 59) },
        { new DateTime(100, 1, 1, 0, 0, 0) },
        { new DateTime(1001, 2, 3, 4, 5, 6) },
        { new DateTime(1899, 12, 29, 23, 59, 59) },
        { new DateTime(1899, 12, 30, 0, 0, 0) },
        { new DateTime(1899, 12, 30, 0, 0, 1) },
        { new DateTime(2000, 1, 2, 1, 2, 3) },
        { new DateTime(2000, 1, 2, 6, 0, 0) },
        { new DateTime(2000, 1, 2, 6, 0, 0, DateTimeKind.Local) },
        { new DateTime(2000, 1, 2, 6, 0, 0, DateTimeKind.Utc) },
        { new DateTime(2025, 6, 2, 0, 0, 0) },
        { new DateTime(2025, 6, 2, 5, 8, 13, 67) },
        { new DateTime(2100, 1, 1, 0, 0, 1) },
        { new DateTime(3000, 1, 2, 6, 0, 0) },
        { DateTime.MaxValue },
    };

    [Theory]
    [MemberData(nameof(DateTimes))]
    public void OADate_TryFormat_Success(DateTime dateTime)
    {
        // Arrange
        var expectedValue = dateTime.ToOADate();
        Span<byte> destination = stackalloc byte[19];
        var oaDate = new OADate(dateTime.Ticks);

        // Act
        var result = oaDate.TryFormat(destination, out var bytesWritten);

        // Assert
        Assert.True(result);
        Assert.True(bytesWritten > 0);

        var bytes = destination.Slice(0, bytesWritten);
        DoublePolyfill.TryParse(bytes, CultureInfo.InvariantCulture, out var actualValue);
        Assert.Equal(expectedValue, actualValue, 0.00000000005);
    }

    [Fact]
    public void OADate_TryFormat_Random()
    {
        // Arrange
        var random = new Random(42);
        var min = new DateTime(2000, 1, 1);
        var max = new DateTime(2100, 1, 1);

        var dateTimes = Enumerable.Repeat(0, 10000)
            .Select(_ => random.NextDateTime(min, max));

        Span<byte> destination = stackalloc byte[19];

        foreach (var dateTime in dateTimes)
        {
            var oaDate = new OADate(dateTime.Ticks);

            // Act
            var result = oaDate.TryFormat(destination, out var bytesWritten);

            // Assert
            Assert.True(result);
            Assert.True(bytesWritten > 0);

            var bytes = destination.Slice(0, bytesWritten);
            DoublePolyfill.TryParse(bytes, CultureInfo.InvariantCulture, out var actualValue);
            Assert.Equal(dateTime.ToOADate(), actualValue, 0.00000000005);
        }
    }
}
#endif