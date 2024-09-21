using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using SpreadCheetah.SourceGenerator.Extensions;
using SpreadCheetah.SourceGenerator.Helpers;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed class PropertyAnalyzer(IDiagnosticsReporter diagnostics)
{
    private PropertyAttributeData _result;

    public PropertyAttributeData Analyze(IPropertySymbol property, CancellationToken token)
    {
        _result = new();

        foreach (var attribute in property.GetAttributes())
        {
            _ = attribute.AttributeClass?.ToDisplayString() switch
            {
                Attributes.CellStyle => TryGetCellStyleAttribute(attribute, token),
                Attributes.CellValueConverter => TryGetCellValueConverterAttribute(attribute, property.Type, token),
                Attributes.CellValueTruncate => TryGetCellValueTruncateAttribute(attribute, property.Type, token),
                Attributes.ColumnHeader => TryGetColumnHeaderAttribute(attribute, token),
                Attributes.ColumnOrder => TryGetColumnOrderAttribute(attribute, token),
                Attributes.ColumnWidth => TryGetColumnWidthAttribute(attribute, token),
                _ => false
            };
        }

        return _result;
    }

    private bool TryGetCellStyleAttribute(
        AttributeData attribute,
        CancellationToken token)
    {
        if (_result.CellStyle is not null)
            return false;

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: string value } arg])
            return false;

        if (string.IsNullOrEmpty(value)
            || value.Length > 255
            || char.IsWhiteSpace(value[0])
            || char.IsWhiteSpace(value[^1]))
        {
            diagnostics.ReportInvalidArgument(attribute, value, token);
            return false;
        }

        _result.CellStyle = new CellStyle(arg.ToCSharpString());
        return true;
    }

    private bool TryGetCellValueConverterAttribute(
        AttributeData attribute,
        ITypeSymbol propertyType,
        CancellationToken token)
    {
        if (_result.CellValueConverter is not null)
            return false;

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: INamedTypeSymbol converterTypeSymbol }])
            return false;

        var typeName = converterTypeSymbol.ToDisplayString();
        var propertyTypeName = propertyType.OriginalDefinition.ToDisplayString();

        if (converterTypeSymbol.BaseType is not { Name: "CellValueConverter", TypeArguments: [INamedTypeSymbol typeArgument] }
            || !string.Equals(typeArgument.OriginalDefinition.ToDisplayString(), propertyTypeName, StringComparison.Ordinal))
        {
            diagnostics.ReportTypeMustInherit(attribute, typeName, $"CellValueConverter<{propertyTypeName}>", token);
            return false;
        }

        var hasPublicConstructor = converterTypeSymbol.Constructors.Any(ctor =>
            ctor is { Parameters.Length: 0, DeclaredAccessibility: Accessibility.Public });

        if (!hasPublicConstructor)
        {
            diagnostics.ReportTypeMustHaveDefaultConstructor(attribute, typeName, token);
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
        if (_result.CellValueTruncate is not null)
            return false;

        if (propertyType.SpecialType != SpecialType.System_String)
        {
            diagnostics.ReportUnsupportedType(attribute, propertyType, token);
            return false;
        }

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: int attributeValue }])
            return false;

        if (attributeValue <= 0)
        {
            diagnostics.ReportInvalidArgument(attribute, attributeValue, token);
            return false;
        }

        _result.CellValueTruncate = new CellValueTruncate(attributeValue);
        return true;
    }

    private bool TryGetColumnHeaderAttribute(
        AttributeData attribute,
        CancellationToken token)
    {
        if (_result.ColumnHeader is not null)
            return false;

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

    private bool TryGetColumnOrderAttribute(
        AttributeData attribute,
        CancellationToken token)
    {
        if (_result.ColumnOrder is not null)
            return false;

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: int attributeValue }])
            return false;

        var location = attribute.GetLocation(token);
        _result.ColumnOrder = new ColumnOrder(attributeValue, location);
        return true;
    }

    private bool TryGetColumnWidthAttribute(
        AttributeData attribute,
        CancellationToken token)
    {
        if (_result.ColumnWidth is not null)
            return false;

        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: double attributeValue }])
            return false;

        if (attributeValue is <= 0 or > 255)
        {
            diagnostics.ReportInvalidArgument(attribute, attributeValue, token);
            return false;
        }

        _result.ColumnWidth = new ColumnWidth(attributeValue);
        return true;
    }
}
