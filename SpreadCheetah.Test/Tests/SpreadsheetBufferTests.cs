using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using SpreadCheetah;
using SpreadCheetah.MetadataXml.Attributes;
using System.Text;

namespace SpreadCheetah.Test.Tests;

public class SpreadsheetBufferTests
{
    private static CancellationToken Token => TestContext.Current.CancellationToken;

    [Fact]
    public async Task SpreadsheetBuffer_FlushToStreamAndContinue_IfBufferSpaceNotEnough()
    {
        // arrange
        using var stream = new MemoryStream();
        using var buffer = new SpreadsheetBuffer(30);

        // act
        var bytesToSkip = 0;
        var writtenBytes = 0;

        buffer.TryWrite("111111111111111"u8);

        var result1 = TryWrite(buffer, bytesToSkip, out writtenBytes);

        bytesToSkip = writtenBytes;
        await buffer.FlushToStreamAsync(stream, Token);

        var result2 = TryWrite(buffer, bytesToSkip, out writtenBytes);

        await buffer.FlushToStreamAsync(stream, Token);

        // assert
        stream.Seek(0, SeekOrigin.Begin);
        using var reader = new StreamReader(stream, Encoding.UTF8);
        var result = await reader.ReadToEndAsync();

        Assert.False(result1);
        Assert.True(result2);
        Assert.Equal("111111111111111123 some_long_attribute_name=\"1\"", result);
    }

    private static bool TryWrite(SpreadsheetBuffer buffer, int bytesToSkip, out int writtenBytes)
    {
        var intAttribute = new IntAttribute("some_long_attribute_name"u8, 1);
        return buffer.TryWrite(bytesToSkip, out writtenBytes,
            $"{"123"u8}" +
            $"{intAttribute}");
    }
}