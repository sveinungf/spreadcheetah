using Microsoft.CodeAnalysis;
using SpreadCheetah.SourceGenerator.Models;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed record ContextClass(
    string Name,
    Accessibility DeclaredAccessibility,
    string? Namespace,
    Dictionary<INamedTypeSymbol, LocationInfo> RowTypes, // TODO: Don't use INamedTypeSymbol
    GeneratorOptions? Options);