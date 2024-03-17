using SpreadCheetah.SourceGenerator.Models;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal static class ColumnHeaderMap
{
    public static ColumnHeaderInfo ToColumnHeaderInfo(this ColumnHeader columnHeader)
    {
        var fullPropertyReference = columnHeader.PropertyReference is { } reference
            ? $"{reference.TypeFullName}.{reference.PropertyName}"
            : null;

        return new ColumnHeaderInfo(
            columnHeader.RawString,
            fullPropertyReference);
    }
}
