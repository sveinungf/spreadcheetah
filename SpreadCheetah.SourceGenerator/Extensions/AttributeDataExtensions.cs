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
    
    public static InheritedColumnsOrderingStrategy? TryGetInheritedColumnOrderingAttribute(this AttributeData attribute)
    {
        if (!string.Equals(Attributes.InheritedColumnOrderStrategy, attribute.AttributeClass?.ToDisplayString(), StringComparison.Ordinal))
            return null;

        return (InheritedColumnsOrderingStrategy)attribute.ConstructorArguments[0].Value!;
    }

    public static ColumnHeader? TryGetColumnHeaderAttribute(this AttributeData attribute, ICollection<DiagnosticInfo> diagnosticInfos, CancellationToken token)
    {
        if (!string.Equals(Attributes.ColumnHeader, attribute.AttributeClass?.ToDisplayString(), StringComparison.Ordinal))
            return null;

        var args = attribute.ConstructorArguments;

        if (args is [{ Value: string } arg])
            return new ColumnHeader(arg.ToCSharpString());

        if (args is [{ Value: INamedTypeSymbol type }, { Value: string propertyName }])
            return TryGetColumnHeaderWithPropertyReference(type, propertyName, attribute, diagnosticInfos, token);

        return null;
    }

    private static ColumnHeader? TryGetColumnHeaderWithPropertyReference(
        INamedTypeSymbol type, string propertyName, AttributeData attribute,
        ICollection<DiagnosticInfo> diagnosticInfos, CancellationToken token)
    {
        var typeFullName = type.ToDisplayString();

        foreach (var member in type.GetMembers())
        {
            if (!string.Equals(member.Name, propertyName, StringComparison.Ordinal))
                continue;

            if (!member.IsStaticPropertyWithPublicGetter(out var p))
                break;

            if (p.Type.SpecialType != SpecialType.System_String)
                break;

            var propertyReference = new ColumnHeaderPropertyReference(typeFullName, propertyName);
            return new ColumnHeader(propertyReference);
        }

        var location = attribute.GetLocation(token);
        diagnosticInfos.Add(new DiagnosticInfo(Diagnostics.InvalidColumnHeaderPropertyReference, location, new([propertyName, typeFullName])));
        return null;
    }

    public static ColumnOrder? TryGetColumnOrderAttribute(this AttributeData attribute, CancellationToken token)
    {
        if (!string.Equals(Attributes.ColumnOrder, attribute.AttributeClass?.ToDisplayString(), StringComparison.Ordinal))
            return null;

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: int attributeValue }])
            return null;

        var location = attribute.GetLocation(token);
        return new ColumnOrder(attributeValue, location);
    }

    private static LocationInfo? GetLocation(this AttributeData attribute, CancellationToken token)
    {
        return attribute
            .ApplicationSyntaxReference?
            .GetSyntax(token)
            .GetLocation()
            .ToLocationInfo();
    }
}
