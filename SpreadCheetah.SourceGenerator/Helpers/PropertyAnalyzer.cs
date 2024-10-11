using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using SpreadCheetah.SourceGenerator.Extensions;
using SpreadCheetah.SourceGenerator.Models;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed class PropertyAnalyzer(IDiagnosticsReporter diagnostics)
{
    private PropertyAttributeData _result;

    public PropertyAttributeData Analyze(IPropertySymbol property, CancellationToken token)
    {
        _result = new();

        foreach (var attribute in property.GetAttributes())
        {
            if (!attribute.HasSrcGenAttributeNamespace())
                continue;

            _ = attribute.AttributeClass?.MetadataName switch
            {
                "CellStyleAttribute" => TryGetCellStyleAttribute(attribute, token),
                "CellValueConverterAttribute" => TryGetCellValueConverterAttribute(attribute, property.Type, token),
                "CellValueConverterAttribute`1" => TryGetCellValueConverterGenericAttribute(attribute, property.Type, token),
                "CellValueTruncateAttribute" => TryGetCellValueTruncateAttribute(attribute, property.Type, token),
                "ColumnHeaderAttribute" => TryGetColumnHeaderAttribute(attribute, token),
                "ColumnOrderAttribute" => TryGetColumnOrderAttribute(attribute),
                "ColumnWidthAttribute" => TryGetColumnWidthAttribute(attribute, token),
                _ => false
            };
        }

        return _result;
    }

    private bool TryGetCellStyleAttribute(
        AttributeData attribute,
        CancellationToken token)
    {
        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: string value } arg])
            return false;

        if (string.IsNullOrEmpty(value)
            || value.Length > 255
            || char.IsWhiteSpace(value[0])
            || char.IsWhiteSpace(value[^1]))
        {
            diagnostics.ReportInvalidArgument(attribute, token);
            return false;
        }

        _result.CellStyle = new CellStyle(arg.ToCSharpString());
        return true;
    }

    private bool TryGetCellValueConverterGenericAttribute(
        AttributeData attribute,
        ITypeSymbol propertyType,
        CancellationToken token)
    {
        var typeArgs = attribute.AttributeClass?.TypeArguments;
        return typeArgs is [INamedTypeSymbol converterType]
            && TryGetCellValueConverterAttribute(attribute, true, converterType, propertyType, token);
    }

    private bool TryGetCellValueConverterAttribute(
        AttributeData attribute,
        ITypeSymbol propertyType,
        CancellationToken token)
    {
        var constructorArgs = attribute.ConstructorArguments;
        return constructorArgs is [{ Value: INamedTypeSymbol converterType }]
            && TryGetCellValueConverterAttribute(attribute, false, converterType, propertyType, token);
    }

    private bool TryGetCellValueConverterAttribute(
        AttributeData attribute,
        bool genericAttribute,
        INamedTypeSymbol converterTypeSymbol,
        ITypeSymbol propertyType,
        CancellationToken token)
    {
        var typeName = converterTypeSymbol.ToDisplayString();
        var propertyTypeName = propertyType.OriginalDefinition.ToDisplayString();

        if (converterTypeSymbol.BaseType is not { Name: "CellValueConverter", TypeArguments: [INamedTypeSymbol typeArgument] }
            || !string.Equals(typeArgument.OriginalDefinition.ToDisplayString(), propertyTypeName, StringComparison.Ordinal))
        {
            diagnostics.ReportTypeMustInherit(attribute, genericAttribute, typeName, $"CellValueConverter<{propertyTypeName}>", token);
            return false;
        }

        var hasPublicConstructor = converterTypeSymbol.Constructors.Any(ctor =>
            ctor is { Parameters.Length: 0, DeclaredAccessibility: Accessibility.Public });

        if (!hasPublicConstructor)
        {
            diagnostics.ReportTypeMustHaveDefaultConstructor(attribute, genericAttribute, typeName, token);
            return false;
        }

        _result.CellValueConverter = new CellValueConverter(typeName);
        return true;
    }

    private bool TryGetCellValueTruncateAttribute(
        AttributeData attribute,
        ITypeSymbol propertyType,
        CancellationToken token)
    {
        if (propertyType.SpecialType != SpecialType.System_String)
        {
            diagnostics.ReportUnsupportedPropertyTypeForAttribute(attribute, propertyType, token);
            return false;
        }

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: int attributeValue }])
            return false;

        if (attributeValue <= 0)
        {
            diagnostics.ReportInvalidArgument(attribute, token);
            return false;
        }

        _result.CellValueTruncate = new CellValueTruncate(attributeValue);
        return true;
    }

    private bool TryGetColumnHeaderAttribute(
        AttributeData attribute,
        CancellationToken token)
    {
        var args = attribute.ConstructorArguments;
        if (args is [{ Value: string } arg])
            _result.ColumnHeader = new ColumnHeader(arg.ToCSharpString());
        else if (args is [{ Value: INamedTypeSymbol type }, { Value: string propertyName }])
            _result.ColumnHeader = TryGetColumnHeaderWithPropertyReference(type, propertyName, attribute, token);

        return _result.ColumnHeader is not null;
    }

    private ColumnHeader? TryGetColumnHeaderWithPropertyReference(
        INamedTypeSymbol type,
        string propertyName,
        AttributeData attribute,
        CancellationToken token)
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

        diagnostics.ReportInvalidPropertyReference(attribute, propertyName, typeFullName, token);
        return null;
    }

    private bool TryGetColumnOrderAttribute(AttributeData attribute)
    {
        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: int attributeValue }])
            return false;

        _result.ColumnOrder = new ColumnOrder(attributeValue);
        return true;
    }

    private bool TryGetColumnWidthAttribute(
        AttributeData attribute,
        CancellationToken token)
    {
        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: double attributeValue }])
            return false;

        if (attributeValue is <= 0 or > 255)
        {
            diagnostics.ReportInvalidArgument(attribute, token);
            return false;
        }

        _result.ColumnWidth = new ColumnWidth(attributeValue);
        return true;
    }
}
