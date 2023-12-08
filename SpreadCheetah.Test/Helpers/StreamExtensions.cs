namespace SpreadCheetah.Test.Helpers;

internal static class StreamExtensions
{
    public static async Task<byte[]> ToByteArrayAsync(this Stream stream)
    {
        if (stream.CanSeek)
            stream.Position = 0;

        if (stream is MemoryStream ms)
            return ms.ToArray();

        using var memoryStream = new MemoryStream();
        await stream.CopyToAsync(memoryStream);
        return memoryStream.ToArray();
    }
}
