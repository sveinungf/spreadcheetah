using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record ColumnHeaderLookupInfo(
    INamedTypeSymbol Type,
    string? Prefix,
    string? Suffix);