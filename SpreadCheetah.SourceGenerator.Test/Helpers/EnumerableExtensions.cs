using Xunit;

namespace SpreadCheetah.SourceGenerator.Test.Helpers;

internal static class EnumerableExtensions
{
    public static TheoryData<T> ToTheoryData<T>(this IEnumerable<T> enumerable)
    {
        var result = new TheoryData<T>();
        foreach (var element in enumerable)
        {
            result.Add(element);
        }

        return result;
    }
}