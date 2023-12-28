using Microsoft.CodeAnalysis;
using SpreadCheetah.SourceGenerator.Models;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed record ContextClass(
    ITypeSymbol ContextClassType,
    Dictionary<INamedTypeSymbol, Location> RowTypes,
    CompilationTypes CompilationTypes,
    GeneratorOptions? Options);