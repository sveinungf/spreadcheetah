using Polyfills;
using SpreadCheetah.Helpers;
using SpreadCheetah.Test.Helpers;
using System.Globalization;

namespace SpreadCheetah.Test.Tests;

public class OADateTests
{
    private readonly DateTime _negativeCutoff = new(1900, 03, 01);

    public static TheoryData<DateTime> DateTimes => new()
    {
        { DateTime.MinValue },
        { new DateTime(1, 1, 1, 23, 59, 59) },
        { new DateTime(100, 1, 1, 0, 0, 0) },
        { new DateTime(1001, 2, 3, 4, 5, 6) },
        { new DateTime(1899, 12, 29, 23, 59, 59) },
        { new DateTime(1899, 12, 30, 0, 0, 0) },
        { new DateTime(1899, 12, 30, 0, 0, 1) },
        { new DateTime(1900, 01, 01) },
        { new DateTime(1900, 02, 28) },
        { new DateTime(1900, 03, 01) },
        { new DateTime(2000, 1, 2, 1, 2, 3) },
        { new DateTime(2000, 1, 2, 6, 0, 0) },
        { new DateTime(2000, 1, 2, 6, 0, 0, DateTimeKind.Local) },
        { new DateTime(2000, 1, 2, 6, 0, 0, DateTimeKind.Utc) },
        { new DateTime(2000, 1, 2, 9, 0, 0) },
        { new DateTime(2000, 1, 2, 12, 0, 0) },
        { new DateTime(2025, 6, 2, 0, 0, 0) },
        { new DateTime(2025, 6, 2, 5, 8, 13, 67) },
        { new DateTime(2100, 1, 1, 0, 0, 1) },
        { new DateTime(3000, 1, 2, 6, 0, 0) },
        { new DateTime(630823680086400000) },
        { new DateTime(630823680008640000) },
        { new DateTime(630823680000864000) },
        { DateTime.MaxValue },
    };

    [Theory]
    [MemberData(nameof(DateTimes))]
    public void OADate_TryFormat_Success(DateTime dateTime)
    {
        // Arrange
        var expectedValue = dateTime.ToOADate();
        if (dateTime < _negativeCutoff && dateTime != DateTime.MinValue)
        {
            // Subtract 1 from integer (day) portion for days before March 1, 1900
            var remainder = expectedValue % 1;
            expectedValue = ((int)expectedValue) - 1;
            if (remainder * expectedValue < 0)
            {
                remainder *= -1;
            }
            expectedValue += remainder;
        }
        Span<byte> destination = stackalloc byte[19];
        var oaDate = new OADate(dateTime.Ticks);

        // Act
        var result = oaDate.TryFormat(destination, out var bytesWritten);

        // Assert
        Assert.True(result);
        Assert.True(bytesWritten > 0);

        var bytes = destination.Slice(0, bytesWritten);
        double.TryParse(bytes, CultureInfo.InvariantCulture, out var actualValue);
        Assert.Equal(expectedValue, actualValue, 0.00000002);
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
            double.TryParse(bytes, CultureInfo.InvariantCulture, out var actualValue);
            Assert.Equal(dateTime.ToOADate(), actualValue, 0.00000002);
        }
    }
}