using System.Collections;
using System.Diagnostics.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Helpers;

/// <summary>
/// Based on the implementation from:
/// https://github.com/CommunityToolkit/dotnet/blob/main/src/CommunityToolkit.Mvvm.SourceGenerators/Helpers/EquatableArray%7BT%7D.cs
/// </summary>
[ExcludeFromCodeCoverage]
internal readonly struct EquatableArray<T>(T[] underlyingArray)
    : IEquatable<EquatableArray<T>>, IReadOnlyCollection<T>
    where T : IEquatable<T>
{
    public static readonly EquatableArray<T> Empty = new([]);

    private readonly T[]? _array = underlyingArray;

    public bool Equals(EquatableArray<T> array)
    {
        return AsSpan().SequenceEqual(array.AsSpan());
    }

    public override bool Equals(object? obj)
    {
        return obj is EquatableArray<T> array && Equals(this, array);
    }

    public override int GetHashCode()
    {
        if (_array is not T[] array)
        {
            return 0;
        }

        HashCode hashCode = default;

        foreach (T item in array)
        {
            hashCode.Add(item);
        }

        return hashCode.ToHashCode();
    }

    public ReadOnlySpan<T> AsSpan() => _array.AsSpan();
    public T[]? GetArray() => _array;

    IEnumerator<T> IEnumerable<T>.GetEnumerator()
    {
        return ((IEnumerable<T>)(_array ?? [])).GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return ((IEnumerable<T>)(_array ?? [])).GetEnumerator();
    }

    public int Count => _array?.Length ?? 0;

    public static bool operator ==(EquatableArray<T> left, EquatableArray<T> right) => left.Equals(right);
    public static bool operator !=(EquatableArray<T> left, EquatableArray<T> right) => !left.Equals(right);
}