using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record CompilationTypes(
    INamedTypeSymbol WorksheetRowAttribute,
    INamedTypeSymbol WorksheetRowGenerationOptionsAttribute,
    INamedTypeSymbol ColumnOrderAttribute,
    INamedTypeSymbol WorksheetRowContext);