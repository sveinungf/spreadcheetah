using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed record TypePropertiesInfo(
    List<string> PropertyNames,
    List<IPropertySymbol> UnsupportedProperties);