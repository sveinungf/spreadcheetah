#if !NET5_0_OR_GREATER

namespace System.Linq;

internal static class EnumerableExtensions
{
    public static IEnumerable<(TFirst First, TSecond Second)> Zip<TFirst, TSecond>(this IEnumerable<TFirst> first, IEnumerable<TSecond> second)
    {
        using IEnumerator<TFirst> e1 = first.GetEnumerator();
        using IEnumerator<TSecond> e2 = second.GetEnumerator();
        while (e1.MoveNext() && e2.MoveNext())
        {
            yield return (e1.Current, e2.Current);
        }
    }
}
#endif