namespace SpreadCheetah.Helpers;

internal static class CollectionExtensions
{
    public static PooledArray<T> ToPooledArray<T>(this ICollection<T> collection)
    {
        return PooledArray<T>.Create(collection);
    }

    public static PooledArray<TTarget> ToPooledArray<TSource, TState, TTarget>(
        this ICollection<TSource> collection,
        TState state,
        Func<TSource, TState, TTarget> convert)
    {
        return PooledArray<TTarget>.Create(collection, state, convert);
    }
}
