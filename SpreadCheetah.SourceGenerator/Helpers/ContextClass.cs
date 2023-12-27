using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed record ContextClass(
    ITypeSymbol ContextClassType,
    Dictionary<INamedTypeSymbol, Location> RowTypes,
    GeneratorOptions? Options);