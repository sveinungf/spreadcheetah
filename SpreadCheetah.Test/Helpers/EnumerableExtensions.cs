using System.Collections.Immutable;

namespace SpreadCheetah.Test.Helpers;

internal static class EnumerableExtensions
{
    public static IList<DataCell> ToIList(this IEnumerable<DataCell> cells, IListType type)
    {
        return type switch
        {
            IListType.Array => cells.ToArray(),
            IListType.List => cells.ToList(),
            IListType.ImmutableList => cells.ToImmutableList(),
            _ => throw new ArgumentOutOfRangeException(nameof(type), type, null)
        };
    }
}
