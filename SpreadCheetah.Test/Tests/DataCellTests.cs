namespace SpreadCheetah.Test.Tests;

public static class DataCellTests
{
    [Fact]
    public static void DataCell_DateTime_InvalidOADate()
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
        var exception = Record.Exception(() => new DataCell(dateTime));

        // Assert
        Assert.IsType<OverflowException>(exception);
        Assert.Equal(expectedMessage, exception.Message);
    }

    [Fact]
    public static void DataCell_NullableDateTime_InvalidOADate()
    {
        // Arrange
        DateTime? dateTime = new DateTime(99, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        string? expectedMessage = null;
        try
        {
            dateTime.GetValueOrDefault().ToOADate();
        }
        catch (OverflowException e)
        {
            expectedMessage = e.Message;
        }

        Assert.NotNull(expectedMessage);

        // Act
        var exception = Record.Exception(() => new DataCell(dateTime));

        // Assert
        Assert.IsType<OverflowException>(exception);
        Assert.Equal(expectedMessage, exception.Message);
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

    [Fact]
    public static void DataCell_DateTime_MinValue()
    {
        // Arrange
        var dateTime = DateTime.MinValue;

        // Act
        var exception = Record.Exception(() => new DataCell(dateTime));

        // Assert
        Assert.Null(exception);
    }

    [Fact]
    public static void DataCell_NullableDateTime_MinValue()
    {
        // Arrange
        DateTime? dateTime = DateTime.MinValue;

        // Act
        var exception = Record.Exception(() => new DataCell(dateTime));

        // Assert
        Assert.Null(exception);
    }
}
