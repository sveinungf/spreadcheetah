namespace SpreadCheetah.Test.Tests;

public static class DataCellTests
{
    [Theory, CombinatorialData]
    public static void DataCell_DateTime_InvalidOADate(bool nullable)
    {
        // Arrange
        var dateTime = new DateTime(99, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        string? expectedMessage = null;
        try
        {
            dateTime.ToOADate();
        }
        catch (OverflowException e)
        {
            expectedMessage = e.Message;
        }

        Assert.NotNull(expectedMessage);

        // Act
        var exception = Record.Exception(() => nullable
            ? new DataCell((DateTime?)dateTime)
            : new DataCell(dateTime));

        // Assert
        Assert.IsType<OverflowException>(exception);
        Assert.Equal(expectedMessage, exception.Message);
    }

    [Theory, CombinatorialData]
    public static void DataCell_DateTime_MinValue(bool nullable)
    {
        // Arrange
        var dateTime = DateTime.MinValue;

        // Act
        var exception = Record.Exception(() => nullable
            ? new DataCell((DateTime?)dateTime)
            : new DataCell(dateTime));

        // Assert
        Assert.Null(exception);
    }

    [Theory, CombinatorialData]
    public static void DataCell_DateTime_MinValueWithTimePart(
        [CombinatorialValues(86399, 86400)] int seconds,
        bool nullable)
    {
        // Arrange
        var dateTime = DateTime.MinValue.AddSeconds(seconds);
        var shouldThrow = seconds == 86400;

        // Act
        var exception = Record.Exception(() => nullable
            ? new DataCell((DateTime?)dateTime)
            : new DataCell(dateTime));

        // Assert
        Assert.Equal(shouldThrow, exception is not null);
    }

    [Fact]
    public static void DataCell_NullableDateTime_NullValue()
    {
        // Arrange
        DateTime? dateTime = null;

        // Act
        var exception = Record.Exception(() => new DataCell(dateTime));

        // Assert
        Assert.Null(exception);
    }
}
