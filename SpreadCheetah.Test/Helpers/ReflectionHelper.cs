using System.Reflection;

namespace SpreadCheetah.Test.Helpers;

internal static class ReflectionHelper
{
    internal static object? GetInstanceField<T>(T instance, string fieldName)
    {
        const BindingFlags bindFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static;
        var field = typeof(T).GetField(fieldName, bindFlags);
        return field is null
            ? throw new ArgumentException("Could not find the field.", nameof(fieldName), null)
            : field.GetValue(instance);
    }
}
