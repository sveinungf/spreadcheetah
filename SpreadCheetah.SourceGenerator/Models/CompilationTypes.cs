using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Models;

internal record CompilationTypes(
    INamedTypeSymbol WorksheetRowAttribute,
    INamedTypeSymbol WorksheetRowGenerationOptionsAttribute,
    INamedTypeSymbol ColumnOrderAttribute,
    INamedTypeSymbol WorksheetRowContext);