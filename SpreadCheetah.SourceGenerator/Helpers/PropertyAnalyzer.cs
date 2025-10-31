using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using SpreadCheetah.SourceGenerator.Extensions;
using SpreadCheetah.SourceGenerator.Models;
using SpreadCheetah.SourceGenerator.Models.Values;
using System.Diagnostics;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed class PropertyAnalyzer(
    IDiagnosticsReporter diagnostics,
    IPropertySymbol propertySymbol)
{
    public PropertyAttributeData Result { get; } = new();

    public void Analyze(CancellationToken token)
    {
        var attributes = propertySymbol
            .GetAttributes()
            .Where(x => x.AttributeClass is { } c && c.HasSpreadCheetahSrcGenNamespace())
            .OrderBy(x => x, AttributeDataComparer.Instance);

        foreach (var attribute in attributes)
        {
            if (Result.ColumnIgnore is not null)
            {
                diagnostics.ReportAttributeCombinationNotSupported(attribute, Attributes.ColumnIgnore, token);
                continue;
            }

            _ = attribute.AttributeClass?.MetadataName switch
            {
                Attributes.CellFormat => TryGetCellFormatAttribute(attribute, propertySymbol.Type, token),
                Attributes.CellStyle => TryGetCellStyleAttribute(attribute, token),
                Attributes.CellValueConverter => TryGetCellValueConverterAttribute(attribute, propertySymbol.Type, token),
                Attributes.CellValueTruncate => TryGetCellValueTruncateAttribute(attribute, propertySymbol.Type, token),
                Attributes.ColumnHeader => TryGetColumnHeaderAttribute(attribute, token),
                Attributes.ColumnIgnore => TryGetColumnIgnoreAttribute(),
                Attributes.ColumnOrder => TryGetColumnOrderAttribute(attribute),
                Attributes.ColumnWidth => TryGetColumnWidthAttribute(attribute, token),
                _ => false
            };
        }
    }

    public void Analyze(InferColumnHeadersInfo inferColumnHeaders)
    {
        if (Result is { ColumnHeader: null, ColumnIgnore: null })
            Result.ColumnHeader = GetColumnHeaderWithPropertyReference(inferColumnHeaders, propertySymbol);
    }

    private bool TryGetCellStyleAttribute(
        AttributeData attribute,
        CancellationToken token)
    {
        Debug.Assert(AttributeDataComparer.Compare(Attributes.CellFormat, Attributes.CellStyle) < 0);

        if (Result.CellFormat is not null)
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

        Result.CellStyle = new CellStyle(arg.ToCSharpString());
        return true;
    }

    private bool TryGetCellFormatAttribute(
        AttributeData attribute,
        ITypeSymbol propertyType,
        CancellationToken token)
    {
        if (!propertyType.SupportsCellFormatAttribute())
        {
            diagnostics.ReportUnsupportedPropertyTypeForAttribute(attribute, propertyType, token);
            return false;
        }

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

            Result.CellFormat = new CellFormat(arg.ToCSharpString());
            return true;
        }

        if (value.IsEnum(out StandardNumberFormat standardFormat))
        {
            Result.CellFormat = new CellFormat(standardFormat);
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

        Result.CellValueConverter = new CellValueConverter(typeName);
        return true;
    }

    private bool TryGetCellValueTruncateAttribute(
        AttributeData attribute,
        ITypeSymbol propertyType,
        CancellationToken token)
    {
        Debug.Assert(AttributeDataComparer.Compare(Attributes.CellValueConverter, Attributes.CellValueTruncate) < 0);

        if (Result.CellValueConverter is not null)
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

        Result.CellValueTruncate = new CellValueTruncate(attributeValue);
        return true;
    }

    private bool TryGetColumnHeaderAttribute(
        AttributeData attribute,
        CancellationToken token)
    {
        var args = attribute.ConstructorArguments;
        if (args is [{ Value: string } arg])
            Result.ColumnHeader = new ColumnHeader(arg.ToCSharpString());
        else if (args is [{ Value: INamedTypeSymbol type }, { Value: string propertyName }])
            Result.ColumnHeader = GetColumnHeaderWithPropertyReference(type, propertyName, attribute, token);

        return Result.ColumnHeader is not null;
    }

    private ColumnHeader? GetColumnHeaderWithPropertyReference(
        INamedTypeSymbol type,
        string propertyName,
        AttributeData attribute,
        CancellationToken token)
    {
        var result = GetColumnHeaderWithPropertyReference(type, propertyName, out var missingReference, out var invalidReference);

        if (missingReference)
            diagnostics.ReportMissingPropertyForColumnHeader(attribute, propertyName, type.ToDisplayString(), token);

        if (invalidReference)
            diagnostics.ReportUnsupportedPropertyForColumnHeader(attribute, propertyName, type.ToDisplayString(), token);

        return result;
    }

    private ColumnHeader? GetColumnHeaderWithPropertyReference(
        InferColumnHeadersInfo inferColumnHeaders,
        IPropertySymbol rowTypeProperty)
    {
        var type = inferColumnHeaders.Type;
        var referencedPropertyName = inferColumnHeaders.Prefix + rowTypeProperty.Name + inferColumnHeaders.Suffix;
        var result = GetColumnHeaderWithPropertyReference(type, referencedPropertyName, out var missingReference, out var invalidReference);

        if (missingReference)
            diagnostics.ReportMissingPropertyForColumnHeader(rowTypeProperty, referencedPropertyName, type.ToDisplayString());

        if (invalidReference)
            diagnostics.ReportUnsupportedPropertyForColumnHeader(rowTypeProperty, referencedPropertyName, type.ToDisplayString());

        return result;
    }

    private static ColumnHeader? GetColumnHeaderWithPropertyReference(
        INamedTypeSymbol type,
        string propertyName,
        out bool missingReference,
        out bool invalidReference)
    {
        missingReference = false;
        invalidReference = false;
        var typeFullName = type.ToDisplayString();

        foreach (var member in type.GetMembers())
        {
            if (!string.Equals(member.Name, propertyName, StringComparison.Ordinal))
                continue;

            if (!member.IsStaticPropertyWithPublicGetter(out var p)
                || p.Type.SpecialType != SpecialType.System_String)
            {
                invalidReference = true;
                return null;
            }

            var propertyReference = new ColumnHeaderPropertyReference(typeFullName, propertyName);
            return new ColumnHeader(propertyReference);
        }

        missingReference = true;
        return null;
    }

    private bool TryGetColumnIgnoreAttribute()
    {
        Result.ColumnIgnore = new ColumnIgnore();
        return true;
    }

    private bool TryGetColumnOrderAttribute(AttributeData attribute)
    {
        var args = attribute.ConstructorArguments;
        if (args is not [{ Value: int attributeValue }])
            return false;

        Result.ColumnOrder = new ColumnOrder(attributeValue);
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

        Result.ColumnWidth = new ColumnWidth(attributeValue);
        return true;
    }
}
