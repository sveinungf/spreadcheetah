namespace SpreadCheetah.Helpers;

internal static class CollectionExtensions
{
    public static PooledArray<T> ToPooledArray<T>(this ICollection<T> collection)
    {
        return PooledArray<T>.Create(collection);
    }
}
