using SpreadCheetah.SourceGenerator.Models;

namespace SpreadCheetah.SourceGenerator.Helpers;
internal sealed record TypePropertiesInfo(
    SortedDictionary<int, ColumnProperty> Properties,
    EquatableArray<string> UnsupportedPropertyTypeNames);