using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using SpreadCheetah.SourceGenerator.Helpers;
using SpreadCheetah.SourceGenerator.Models;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;

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

    public static InheritedColumnOrder? TryGetInheritedColumnOrderingAttribute(this AttributeData attribute)
    {
        if (!string.Equals(Attributes.InheritColumns, attribute.AttributeClass?.ToDisplayString(), StringComparison.Ordinal))
            return null;

        if (attribute.NamedArguments is [{ Value.Value: { } arg }] && Enum.IsDefined(typeof(InheritedColumnOrder), arg))
            return (InheritedColumnOrder)arg;

        return InheritedColumnOrder.InheritedColumnsFirst;
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

    public static ColumnWidth? TryGetColumnWidthAttribute(this AttributeData attribute,
        ICollection<DiagnosticInfo> diagnosticInfos, CancellationToken token)
    {
        if (!string.Equals(Attributes.ColumnWidth, attribute.AttributeClass?.ToDisplayString(), StringComparison.Ordinal))
            return null;

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: double attributeValue }])
            return null;

        if (attributeValue is <= 0 or > 255)
        {
            var location = attribute.GetLocation(token);
            var stringValue = attributeValue.ToString(CultureInfo.InvariantCulture);
            diagnosticInfos.Add(new DiagnosticInfo(Diagnostics.InvalidAttributeArgument, location, new([stringValue, Attributes.ColumnWidth])));
            return null;
        }

        return new ColumnWidth(attributeValue);
    }

    public static CellValueTruncate? TryGetCellValueTruncateAttribute(this AttributeData attribute, ITypeSymbol propertyType,
        ICollection<DiagnosticInfo> diagnosticInfos, CancellationToken token)
    {
        if (!string.Equals(Attributes.CellValueTruncate, attribute.AttributeClass?.ToDisplayString(), StringComparison.Ordinal))
            return null;

        if (propertyType.SpecialType != SpecialType.System_String)
        {
            var location = attribute.GetLocation(token);
            var typeFullName = propertyType.ToDisplayString();
            diagnosticInfos.Add(new DiagnosticInfo(Diagnostics.UnsupportedTypeForCellValueLengthLimit, location, new([typeFullName])));
            return null;
        }

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: int attributeValue }])
            return null;

        if (attributeValue <= 0)
        {
            var location = attribute.GetLocation(token);
            var stringValue = attributeValue.ToString(CultureInfo.InvariantCulture);
            diagnosticInfos.Add(new DiagnosticInfo(Diagnostics.InvalidAttributeArgument, location, new([stringValue, Attributes.CellValueTruncate])));
            return null;
        }

        return new CellValueTruncate(attributeValue);
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
