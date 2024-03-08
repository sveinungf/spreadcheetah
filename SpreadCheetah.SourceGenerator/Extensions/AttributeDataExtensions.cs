using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using SpreadCheetah.SourceGenerator.Helpers;
using SpreadCheetah.SourceGenerator.Models;
using System.Diagnostics.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Extensions;

internal static class AttributeDataExtensions
{
    public static bool TryParseWorksheetRowAttribute(
        this AttributeData attribute,
        CancellationToken token,
        [NotNullWhen(true)] out INamedTypeSymbol? typeSymbol,
        [NotNullWhen(true)] out Location? location)
    {
        typeSymbol = null;
        location = null;

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: INamedTypeSymbol symbol }])
            return false;

        if (symbol.Kind == SymbolKind.ErrorType)
            return false;

        var syntaxReference = attribute.ApplicationSyntaxReference;
        if (syntaxReference is null)
            return false;

        location = syntaxReference.GetSyntax(token).GetLocation();
        typeSymbol = symbol;
        return true;
    }

    public static bool TryParseOptionsAttribute(
        this AttributeData attribute,
        [NotNullWhen(true)] out GeneratorOptions? options)
    {
        options = null;

        if (!string.Equals(Attributes.GenerationOptions, attribute.AttributeClass?.ToDisplayString(), StringComparison.Ordinal))
            return false;

        if (attribute.NamedArguments.IsDefaultOrEmpty)
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

    public static ColumnHeader? TryGetColumnHeaderAttribute(this AttributeData attribute)
    {
        if (!string.Equals(Attributes.ColumnHeader, attribute.AttributeClass?.ToDisplayString(), StringComparison.Ordinal))
            return null;

        var args = attribute.ConstructorArguments;

        // TODO: Test for when the arguments are in the opposite order (by using named arguments)
        // TODO: Check that "type" has a public static property called "propertyName"

        return args switch
        {
            [{ Value: string } arg] => new ColumnHeader(arg.ToCSharpString()),
            [{ Value: Type type }, { Value: string propertyName }] => new ColumnHeader(type, propertyName),
            _ => null
        };
    }

    public static ColumnOrder? TryGetColumnOrderAttribute(this AttributeData attribute, CancellationToken token)
    {
        if (!string.Equals(Attributes.ColumnOrder, attribute.AttributeClass?.ToDisplayString(), StringComparison.Ordinal))
            return null;

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: int attributeValue }])
            return null;

        var location = attribute
            .ApplicationSyntaxReference?
            .GetSyntax(token)
            .GetLocation()
            .ToLocationInfo();

        return new ColumnOrder(attributeValue, location);
    }
}
