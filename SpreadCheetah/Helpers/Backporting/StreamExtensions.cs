#if NETSTANDARD2_0
namespace System.IO;

internal static class StreamExtensions
{
    public static ValueTask DisposeAsync(this Stream stream)
    {
        stream.Dispose();
        return default;
    }
}
#endif
