#if NETSTANDARD2_0
namespace System.IO;

internal static class DisposableExtensions
{
    public static ValueTask DisposeAsync(this IDisposable disposable)
    {
        disposable.Dispose();
        return default;
    }
}
#endif
