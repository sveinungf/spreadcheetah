using SpreadCheetah.Validations;

namespace SpreadCheetah.Test.Tests;

public class DataValidationTests
{
    [Theory, CombinatorialData]
    public void DataValidation_ListValues_InvalidListValues(bool withComma, bool tooLong)
    {
        // Arrange
        List<string> values = ["One", "\"Two\"", "<Three>"];

        if (withComma)
            values.Add("Fo,ur");
        if (tooLong)
            values.Add(new string('x', 250));

        // Act
        var exception = Record.Exception(() => DataValidation.ListValues(values));

        // Assert
        Assert.Equal(withComma || tooLong, exception is not null);
    }

    [Theory, CombinatorialData]
    public void DataValidation_TryCreateListValues_InvalidListValues(bool withComma, bool tooLong)
    {
        // Arrange
        List<string> values = ["One", "\"Two\"", "<Three>"];

        if (withComma)
            values.Add("Fo,ur");
        if (tooLong)
            values.Add(new string('x', 250));

        // Act
        var ok = DataValidation.TryCreateListValues(values, true, out _);

        // Assert
        Assert.Equal(withComma || tooLong, !ok);
    }

    [Fact]
    public void DataValidation_DateTimeBetween_MinGreaterThanMax()
    {
        // Arrange
        var min = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        var max = min.AddDays(-1);

        // Act & Assert
        Assert.Throws<ArgumentException>(() => DataValidation.DateTimeBetween(min, max));
    }

    [Fact]
    public void DataValidation_DecimalBetween_MinGreaterThanMax()
    {
        // Arrange
        const double min = 10.0;
        const double max = 0.0;

        // Act & Assert
        Assert.Throws<ArgumentException>(() => DataValidation.DecimalBetween(min, max));
    }

    [Fact]
    public void DataValidation_IntegerBetween_MinGreaterThanMax()
    {
        // Arrange
        const int min = 10;
        const int max = 0;

        // Act & Assert
        Assert.Throws<ArgumentException>(() => DataValidation.IntegerBetween(min, max));
    }

    [Fact]
    public void DataValidation_TextLengthBetween_MinGreaterThanMax()
    {
        // Arrange
        const int min = 10;
        const int max = 0;

        // Act & Assert
        Assert.Throws<ArgumentException>(() => DataValidation.TextLengthBetween(min, max));
    }

    [Fact]
    public void DataValidation_ErrorType_SetInvalidType()
    {
        // Arrange
        var dataValidation = DataValidation.TextLengthBetween(0, 10);
        var errorType = (ValidationErrorType)3;

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() => dataValidation.ErrorType = errorType);
    }
}
