using System.Buffers;

namespace SpreadCheetah.Helpers;

internal readonly struct PooledArray<T> : IDisposable
{
    private readonly T[] _array;
    private readonly int _count;

    public ReadOnlyMemory<T> Memory => _array.AsMemory(0, _count);

    public PooledArray()
    {
        _array = Array.Empty<T>();
        _count = 0;
    }

    private PooledArray(T[] array, int count)
    {
        _array = array;
        _count = count;
    }

    public static PooledArray<T> Create(ICollection<T> collection)
    {
        if (collection.Count == 0)
            return new PooledArray<T>();

        var array = ArrayPool<T>.Shared.Rent(collection.Count);

        try
        {
            collection.CopyTo(array, 0);
        }
        catch (Exception)
        {
            ArrayPool<T>.Shared.Return(array);
            throw;
        }

        return new PooledArray<T>(array, collection.Count);
    }

    public void Dispose()
    {
        if (_array is { Length: > 0 })
            ArrayPool<T>.Shared.Return(_array, true);
    }
}
