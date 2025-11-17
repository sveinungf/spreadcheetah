#if NETSTANDARD2_0
namespace System.Threading.Tasks;

internal static class ValueTaskExtensions
{
    extension(ValueTask)
    {
        public static ValueTask CompletedTask => default;
    }
}
#endif