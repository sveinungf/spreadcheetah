using SpreadCheetah.SourceGenerator.Models;

namespace SpreadCheetah.SourceGenerator.Helpers;
internal sealed record TypePropertiesInfo(
    SortedDictionary<int, RowTypeProperty> Properties,
    EquatableArray<string> UnsupportedPropertyTypeNames);