using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed record ContextClass(
    string Name,
    Accessibility DeclaredAccessibility,
    string? Namespace,
    Dictionary<INamedTypeSymbol, Location> RowTypes, // TODO: Don't use INamedTypeSymbol
    GeneratorOptions? Options);