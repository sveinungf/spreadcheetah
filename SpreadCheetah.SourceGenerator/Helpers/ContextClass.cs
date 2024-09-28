using Microsoft.CodeAnalysis;
using SpreadCheetah.SourceGenerator.Models;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed record ContextClass(
    string Name,
    Accessibility DeclaredAccessibility,
    string? Namespace,
    EquatableArray<RowType> RowTypes);