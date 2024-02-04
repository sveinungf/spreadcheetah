using Microsoft.CodeAnalysis;
using SpreadCheetah.SourceGenerator.Models;

namespace SpreadCheetah.SourceGenerator.Helpers;
internal sealed record TypePropertiesInfo(
    SortedDictionary<int, ColumnProperty> Properties,
    List<IPropertySymbol> UnsupportedProperties);