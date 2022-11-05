using SpreadCheetah.Validations;
using Xunit;

namespace SpreadCheetah.Test.Tests;

public class DataValidationTests
{
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void DataValidation_ListValues_InvalidListValues(bool withComma, bool tooLong)
    {
        // Arrange
        var values = new List<string> { "One", "\"Two\"", "<Three>" };

        if (withComma)
            values.Add("Fo,ur");
        if (tooLong)
            values.Add(new string('x', 250));

        // Act
        var exception = Record.Exception(() => DataValidation.ListValues(values));

        // Assert
        Assert.Equal(withComma || tooLong, exception is not null);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void DataValidation_TryCreateListValues_InvalidListValues(bool withComma, bool tooLong)
    {
        // Arrange
        var values = new List<string> { "One", "\"Two\"", "<Three>" };

        if (withComma)
            values.Add("Fo,ur");
        if (tooLong)
            values.Add(new string('x', 250));

        // Act
        var ok = DataValidation.TryCreateListValues(values, true, out _);

        // Assert
        Assert.Equal(withComma || tooLong, !ok);
    }
}
