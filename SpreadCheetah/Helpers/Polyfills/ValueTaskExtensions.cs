#if NETSTANDARD2_0
namespace System.Threading.Tasks;

internal static class ValueTaskExtensions
{
    extension(ValueTask)
    {
        public static ValueTask CompletedTask => default;

#pragma warning disable VSTHRD200 // Use "Async" suffix for async methods
        public static ValueTask<T> FromResult<T>(T result) => new(result);
#pragma warning restore VSTHRD200 // Use "Async" suffix for async methods
    }
}
#endif