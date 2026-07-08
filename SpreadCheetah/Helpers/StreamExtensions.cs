namespace SpreadCheetah.Helpers;

internal static class StreamExtensions
{
    public static ValueTask<PooledArray<byte>> ReadToPooledArrayAsync(this Stream stream, int maxBytesToRead, CancellationToken token)
    {
        return PooledArray<byte>.CreateFromStreamAsync(stream, maxBytesToRead, token);
    }
}
