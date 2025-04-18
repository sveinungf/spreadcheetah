#if !NET7_0_OR_GREATER
namespace SpreadCheetah.Test.Helpers;

internal static class StreamReaderExtensions
{
    public static Task<string> ReadToEndAsync(this StreamReader reader, CancellationToken token)
    {
        _ = token;
#pragma warning disable MA0040 // Forward the CancellationToken parameter to methods that take one
        return reader.ReadToEndAsync();
#pragma warning restore MA0040 // Forward the CancellationToken parameter to methods that take one
    }
}
#endif