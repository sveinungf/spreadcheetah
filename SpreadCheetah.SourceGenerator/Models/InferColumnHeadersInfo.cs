using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record InferColumnHeadersInfo(
    INamedTypeSymbol Type,
    string? Prefix,
    string? Suffix);