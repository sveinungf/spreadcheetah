using Microsoft.CodeAnalysis;
using SpreadCheetah.SourceGenerator.Models;
using System.Diagnostics.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Extensions;

internal static class CompilationExtensions
{
    [ExcludeFromCodeCoverage]
    public static bool TryGetCompilationTypes(
        this Compilation compilation,
        [NotNullWhen(true)] out CompilationTypes? result)
    {
        result = null;
        const string ns = "SpreadCheetah.SourceGeneration";

        if (!compilation.TryGetType($"{ns}.ColumnOrderAttribute", out var columnOrder))
            return false;
        if (!compilation.TryGetType($"{ns}.WorksheetRowContext", out var context))
            return false;
        if (!compilation.TryGetType($"{ns}.WorksheetRowGenerationOptionsAttribute", out var options))
            return false;

        result = new CompilationTypes(
            ColumnOrderAttribute: columnOrder,
            WorksheetRowContext: context,
            WorksheetRowGenerationOptionsAttribute: options);

        return true;
    }

    private static bool TryGetType(
        this Compilation compilation,
        string fullyQualifiedMetadataName,
        [NotNullWhen(true)] out INamedTypeSymbol? result)
    {
        result = compilation.GetTypeByMetadataName(fullyQualifiedMetadataName);
        return result is not null;
    }
}
