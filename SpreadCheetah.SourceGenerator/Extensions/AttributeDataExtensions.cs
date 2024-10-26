using Microsoft.CodeAnalysis;
using SpreadCheetah.SourceGenerator.Helpers;
using System.Diagnostics.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Extensions;

internal static class AttributeDataExtensions
{
    public static bool TryParseWorksheetRowAttribute(
        this AttributeData attribute,
        [NotNullWhen(true)] out INamedTypeSymbol? typeSymbol)
    {
        typeSymbol = null;

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: INamedTypeSymbol symbol }])
            return false;

        if (symbol.Kind == SymbolKind.ErrorType)
            return false;

        typeSymbol = symbol;
        return true;
    }

    public static bool TryParseOptionsAttribute(
        this AttributeData attribute,
        [NotNullWhen(true)] out GeneratorOptions? options)
    {
        options = null;

        if (attribute is not { AttributeClass.MetadataName: Attributes.GenerationOptions })
            return false;
        if (!attribute.AttributeClass.HasSpreadCheetahSrcGenNamespace())
            return false;

        foreach (var (key, value) in attribute.NamedArguments)
        {
            if (!string.Equals(key, "SuppressWarnings", StringComparison.Ordinal))
                continue;

            if (value.Value is bool suppressWarnings)
            {
                options = new GeneratorOptions(suppressWarnings);
                return true;
            }
        }

        return false;
    }
}
