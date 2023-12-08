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
            return new PooledArray<T>(array, collection.Count);
        }
        catch
        {
            ArrayPool<T>.Shared.Return(array);
            throw;
        }
    }

    public static async ValueTask<PooledArray<byte>> CreateAsync(Stream stream, int maxBytesToRead, CancellationToken token)
    {
        var array = ArrayPool<byte>.Shared.Rent(maxBytesToRead);

        try
        {
            var bytesRead = await stream.ReadAsync(array, token).ConfigureAwait(false);
            return new PooledArray<byte>(array, bytesRead);
        }
        catch
        {
            ArrayPool<byte>.Shared.Return(array);
            throw;
        }
    }

    public void Dispose()
    {
        if (_array is { Length: > 0 })
            ArrayPool<T>.Shared.Return(_array, true);
    }
}
