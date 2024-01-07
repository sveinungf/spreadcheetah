using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed record TypePropertiesInfo(
    SortedDictionary<int, string> PropertyNames,
    List<IPropertySymbol> UnsupportedProperties);