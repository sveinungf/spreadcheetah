using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record CompilationTypes(
    INamedTypeSymbol WorksheetRowGenerationOptionsAttribute,
    INamedTypeSymbol ColumnOrderAttribute,
    INamedTypeSymbol WorksheetRowContext);