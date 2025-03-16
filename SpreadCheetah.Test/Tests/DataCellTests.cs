namespace SpreadCheetah.Test.Tests;

public static class DataCellTests
{
    [Fact]
    public static void DataCell_DateTime_InvalidOADate()
    {
        // Arrange
        var dateTime = new DateTime(99, 1, 1);
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
}
