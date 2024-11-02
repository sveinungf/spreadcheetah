using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using SpreadCheetah.SourceGenerator.Extensions;
using SpreadCheetah.SourceGenerator.Models;
using SpreadCheetah.SourceGenerator.Models.Values;
using System.Diagnostics;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed class PropertyAnalyzer(IDiagnosticsReporter diagnostics)
{
    private static AttributeDataComparer AttributeDataComparer { get; } = new();

    private PropertyAttributeData _result;

    public PropertyAttributeData Analyze(IPropertySymbol property, CancellationToken token)
    {
        _result = new();

        var attributes = property
            .GetAttributes()
            .Where(x => x.AttributeClass is { } c && c.HasSpreadCheetahSrcGenNamespace())
            .OrderBy(x => x, AttributeDataComparer);

        foreach (var attribute in attributes)
        {
            if (_result.ColumnIgnore is not null)
            {
                diagnostics.ReportAttributeCombinationNotSupported(attribute, Attributes.ColumnIgnore, token);
                continue;
            }

            _ = attribute.AttributeClass?.MetadataName switch
            {
                Attributes.CellFormat => TryGetCellFormatAttribute(attribute, token),
                Attributes.CellStyle => TryGetCellStyleAttribute(attribute, token),
                Attributes.CellValueConverter => TryGetCellValueConverterAttribute(attribute, property.Type, token),
                Attributes.CellValueTruncate => TryGetCellValueTruncateAttribute(attribute, property.Type, token),
                Attributes.ColumnHeader => TryGetColumnHeaderAttribute(attribute, token),
                Attributes.ColumnIgnore => TryGetColumnIgnoreAttribute(),
                Attributes.ColumnOrder => TryGetColumnOrderAttribute(attribute),
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
        Debug.Assert(AttributeDataComparer.Compare(Attributes.CellFormat, Attributes.CellStyle) < 0);

        if (_result.CellFormat is not null)
        {
            diagnostics.ReportAttributeCombinationNotSupported(attribute, Attributes.CellFormat, token);
            return false;
        }

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

    private bool TryGetCellFormatAttribute(
        AttributeData attribute,
        CancellationToken token)
    {
        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: { } value } arg])
            return false;

        if (value is string stringValue)
        {
            if (stringValue.Length > 255)
            {
                diagnostics.ReportInvalidArgument(attribute, token);
                return false;
            }

            _result.CellFormat = new CellFormat(arg.ToCSharpString());
            return true;
        }

        if (value.IsEnum(out StandardNumberFormat standardFormat))
        {
            _result.CellFormat = new CellFormat(standardFormat);
            return true;
        }

        diagnostics.ReportInvalidArgument(attribute, token);
        return false;
    }

    private bool TryGetCellValueConverterAttribute(
        AttributeData attribute,
        ITypeSymbol propertyType,
        CancellationToken token)
    {
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
        Debug.Assert(AttributeDataComparer.Compare(Attributes.CellValueConverter, Attributes.CellValueTruncate) < 0);

        if (_result.CellValueConverter is not null)
        {
            diagnostics.ReportAttributeCombinationNotSupported(attribute, Attributes.CellValueConverter, token);
            return false;
        }

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

    private bool TryGetColumnIgnoreAttribute()
    {
        _result.ColumnIgnore = new ColumnIgnore();
        return true;
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
