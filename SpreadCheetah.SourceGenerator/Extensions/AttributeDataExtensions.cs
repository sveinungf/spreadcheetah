using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using SpreadCheetah.SourceGenerator.Helpers;
using SpreadCheetah.SourceGenerator.Models;
using System.Collections.Immutable;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace SpreadCheetah.SourceGenerator.Extensions;

internal static class AttributeDataExtensions
{
    public static PropertyAttributeData MapToPropertyAttributeData(
        this ImmutableArray<AttributeData> attributes,
        ITypeSymbol propertyType,
        List<DiagnosticInfo> diagnosticInfos,
        CancellationToken token)
    {
        var result = new PropertyAttributeData();

        foreach (var attribute in attributes)
        {
            var displayString = attribute.AttributeClass?.ToDisplayString();
            if (displayString is null)
                continue;

            _ = displayString switch
            {
                Attributes.CellStyle => attribute.TryGetCellStyleAttribute(diagnosticInfos, token, ref result),
                Attributes.CellValueTruncate => attribute.TryGetCellValueTruncateAttribute(propertyType, diagnosticInfos, token, ref result),
                Attributes.ColumnHeader => attribute.TryGetColumnHeaderAttribute(diagnosticInfos, token, ref result),
                Attributes.ColumnOrder => attribute.TryGetColumnOrderAttribute(token, ref result),
                Attributes.ColumnWidth => attribute.TryGetColumnWidthAttribute(diagnosticInfos, token, ref result),
                Attributes.CellValueConverter => attribute.TryGetCellValueConverterAttribute(propertyType, diagnosticInfos, token, ref result),
                _ => false
            };
        }

        return result;
    }

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

    private static bool TryGetColumnHeaderAttribute(this AttributeData attribute,
        List<DiagnosticInfo> diagnosticInfos, CancellationToken token, ref PropertyAttributeData result)
    {
        if (result.ColumnHeader is not null)
            return false;

        var args = attribute.ConstructorArguments;
        if (args is [{ Value: string } arg])
            result.ColumnHeader = new ColumnHeader(arg.ToCSharpString());
        else if (args is [{ Value: INamedTypeSymbol type }, { Value: string propertyName }])
            result.ColumnHeader = TryGetColumnHeaderWithPropertyReference(type, propertyName, attribute, diagnosticInfos, token);

        return result.ColumnHeader is not null;
    }

    private static ColumnHeader? TryGetColumnHeaderWithPropertyReference(
        INamedTypeSymbol type, string propertyName, AttributeData attribute,
        List<DiagnosticInfo> diagnosticInfos, CancellationToken token)
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
        diagnosticInfos.Add(Diagnostics.InvalidColumnHeaderPropertyReference(location, propertyName, typeFullName));
        return null;
    }

    private static bool TryGetColumnOrderAttribute(this AttributeData attribute, CancellationToken token, ref PropertyAttributeData result)
    {
        if (result.ColumnOrder is not null)
            return false;

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: int attributeValue }])
            return false;

        var location = attribute.GetLocation(token);
        result.ColumnOrder = new ColumnOrder(attributeValue, location);
        return true;
    }

    private static bool TryGetCellStyleAttribute(this AttributeData attribute,
        List<DiagnosticInfo> diagnosticInfos, CancellationToken token, ref PropertyAttributeData result)
    {
        if (result.CellStyle is not null)
            return false;

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: string value } arg])
            return false;

        if (string.IsNullOrEmpty(value)
            || value.Length > 255
            || char.IsWhiteSpace(value[0])
            || char.IsWhiteSpace(value[^1]))
        {
            diagnosticInfos.Add(Diagnostics.InvalidAttributeArgument(attribute, value, token));
            return false;
        }

        result.CellStyle = new CellStyle(arg.ToCSharpString());
        return true;
    }

    private static bool TryGetColumnWidthAttribute(this AttributeData attribute,
        List<DiagnosticInfo> diagnosticInfos, CancellationToken token, ref PropertyAttributeData result)
    {
        if (result.ColumnWidth is not null)
            return false;

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: double attributeValue }])
            return false;

        if (attributeValue is <= 0 or > 255)
        {
            var stringValue = attributeValue.ToString(CultureInfo.InvariantCulture);
            diagnosticInfos.Add(Diagnostics.InvalidAttributeArgument(attribute, stringValue, token));
            return false;
        }

        result.ColumnWidth = new ColumnWidth(attributeValue);
        return true;
    }

    private static bool TryGetCellValueTruncateAttribute(this AttributeData attribute, ITypeSymbol propertyType,
        List<DiagnosticInfo> diagnosticInfos, CancellationToken token, ref PropertyAttributeData data)
    {
        if (data.CellValueTruncate is not null)
            return false;

        if (propertyType.SpecialType != SpecialType.System_String)
        {
            var typeFullName = propertyType.ToDisplayString();
            diagnosticInfos.Add(Diagnostics.UnsupportedTypeForAttribute(attribute, typeFullName, token));
            return false;
        }

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: int attributeValue }])
            return false;

        if (attributeValue <= 0)
        {
            var stringValue = attributeValue.ToString(CultureInfo.InvariantCulture);
            diagnosticInfos.Add(Diagnostics.InvalidAttributeArgument(attribute, stringValue, token));
            return false;
        }

        data.CellValueTruncate = new CellValueTruncate(attributeValue);
        return true;
    }

    private static bool TryGetCellValueConverterAttribute(this AttributeData attribute,
        ITypeSymbol propertyType,
        List<DiagnosticInfo> diagnosticInfos, CancellationToken token,
        ref PropertyAttributeData data)
    {
        var args = attribute.ConstructorArguments;
        if (args.Length == 0)
            return false;

        var cellValueConverterTypeSymbol = args[0].Value as INamedTypeSymbol;
        var cellValueConverterType = cellValueConverterTypeSymbol!.BaseType;
        var typeName = cellValueConverterTypeSymbol.ToDisplayString();

        if (cellValueConverterType is null || !string.Equals(cellValueConverterType.Name, "CellValueConverter", StringComparison.Ordinal))
        {
            diagnosticInfos.Add(Diagnostics.AttributeTypeArgumentMustInherit(attribute, typeName, "CellValueConverter<T>", token));
            return false;
        }

        var publicConstructor = cellValueConverterTypeSymbol.Constructors.FirstOrDefault(symbol =>
            symbol.Parameters.Length == 0 && symbol.DeclaredAccessibility == Accessibility.Public);

        if (publicConstructor is null)
        {
            var errorLocation = attribute.GetLocation(token);
            diagnosticInfos.Add(new DiagnosticInfo(Diagnostics.CellValueConverterWithoutPublicParameterlessConstructor,
                errorLocation, new([typeName, Attributes.CellValueConverter])));
            return false;
        }

        var cellValueMapperInterfaceGenericArgument = cellValueConverterType.TypeArguments[0] as INamedTypeSymbol;

        var isPropertyTypeAndCellValueConverterTypeEquals = string.Equals(
            cellValueMapperInterfaceGenericArgument!.OriginalDefinition.ToDisplayString(),
            propertyType.OriginalDefinition.ToDisplayString(), StringComparison.Ordinal);

        if (!isPropertyTypeAndCellValueConverterTypeEquals)
        {
            var errorLocation = attribute.GetLocation(token);
            diagnosticInfos.Add(new DiagnosticInfo(Diagnostics.CellValueConverterArgumentTypeNotSameAsPropertyType,
                errorLocation, new([typeName, Attributes.CellValueConverter])));
            return false;
        }

        data.CellValueConverter = new CellValueConverter(typeName);
        return true;
    }

    public static LocationInfo? GetLocation(this AttributeData attribute, CancellationToken token)
    {
        return attribute
            .ApplicationSyntaxReference?
            .GetSyntax(token)
            .GetLocation()
            .ToLocationInfo();
    }
}
