using System.Reflection;

namespace SpreadCheetah.SourceGenerator.Test.Helpers;

internal static class TypeExtensions
{
    public static Dictionary<string, PropertyInfo> ToPropertyDictionary(this Type type)
    {
        return type
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .ToDictionary(x => x.Name, StringComparer.Ordinal);
    }
}
