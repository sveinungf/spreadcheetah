using SpreadCheetah.MetadataXml.Attributes;
using SpreadCheetah.Styling;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetBufferTests
{
    // NOTE: 16 is minimal length
    [Theory]
    [InlineData(16, 16)] // No space for ' '
    [InlineData(14, 16)] // No space for attribute name
    [InlineData(10, 16)] // No space for '=\'
    [InlineData(8, 16)] // No space for value
    [InlineData(6, 16)] // No space for '"'
    public void IntAttribute_BufferSizeIsNotEnough(int bufferInitialPosition, int bufferSize)
    {
        // Arrange
        using var buffer = new SpreadsheetBuffer(bufferSize);
        buffer.Advance(bufferInitialPosition);
        var attribute = new IntAttribute("name"u8, 100);

        // Act 
        var result = buffer.TryWrite($"{attribute}");

        // Assert
        Assert.False(result);
    }

    // NOTE: 16 is minimal length
    [Theory]
    [InlineData(16, 16)] // No space for ' '
    [InlineData(14, 16)] // No space for attribute name
    [InlineData(10, 16)] // No space for '=\'
    [InlineData(9, 16)] // No space for value
    [InlineData(8, 16)] // No space for '"'
    public void BooleanAttribute_BufferSizeIsNotEnough(int bufferInitialPosition, int bufferSize)
    {
        // Arrange
        using var buffer = new SpreadsheetBuffer(bufferSize);
        buffer.Advance(bufferInitialPosition);
        var attribute = new BooleanAttribute("name"u8, false);

        // Act 
        var result = buffer.TryWrite($"{attribute}");

        // Assert
        Assert.False(result);
    }

    // NOTE: 16 is minimal length
    [Theory]
    [InlineData(16, 16)] // No space for ' '
    [InlineData(14, 16)] // No space for attribute name
    [InlineData(10, 16)] // No space for '=\'
    [InlineData(9, 16)] // No space for value
    [InlineData(4, 16)] // No space for '"'
    public void SpanByteAttribute_BufferSizeIsNotEnough(int bufferInitialPosition, int bufferSize)
    {
        // Arrange
        using var buffer = new SpreadsheetBuffer(bufferSize);
        buffer.Advance(bufferInitialPosition);
        var attribute = new SpanByteAttribute("name"u8, "value"u8);

        // Act 
        var result = buffer.TryWrite($"{attribute}");

        // Assert
        Assert.False(result);
    }

    // NOTE: 16 is minimal length
    [Theory]
    [InlineData(16, 16)] // No space for ' '
    [InlineData(14, 16)] // No space for attribute name
    [InlineData(10, 16)] // No space for '=\'
    [InlineData(8, 16)] // No space for value
    [InlineData(7, 16)] // No space for '"'
    public void SimpleSingleCellReferenceAttribute_BufferSizeIsNotEnough(int bufferInitialPosition, int bufferSize)
    {
        // Arrange
        using var buffer = new SpreadsheetBuffer(bufferSize);
        buffer.Advance(bufferInitialPosition);
        var attribute = new SimpleSingleCellReferenceAttribute("name"u8, new CellReferences.SimpleSingleCellReference(1, 2));

        // Act 
        var result = buffer.TryWrite($"{attribute}");

        // Assert
        Assert.False(result);
    }
}