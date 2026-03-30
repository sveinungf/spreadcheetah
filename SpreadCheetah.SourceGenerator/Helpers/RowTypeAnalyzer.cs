using Microsoft.CodeAnalysis;
using SpreadCheetah.SourceGenerator.Extensions;
using SpreadCheetah.SourceGenerator.Models;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed class RowTypeAnalyzer(
    IDiagnosticsReporter diagnostics,
    INamedTypeSymbol rowType)
{
    public RowTypeAttributeData Result { get; } = new();

    public void Analyze(CancellationToken token)
    {
        foreach (var attribute in GetAllAttributes(rowType))
        {
            _ = attribute.AttributeClass?.MetadataName switch
            {
                Attributes.DefaultColumnWidth => TryGetDefaultColumnWidthAttribute(attribute, token),
                _ => false
            };
        }
    }

    private static List<AttributeData> GetAllAttributes(INamedTypeSymbol type)
    {
        var attributes = type.GetAttributes();
        if (attributes.Length == 0)
            return [];

        var spreadCheetahAttributes = attributes
            .Where(x => x.AttributeClass is { HasSpreadCheetahSrcGenNamespace: true })
            .ToList();

        if (spreadCheetahAttributes.Count == 0)
            return [];

        var inheritColumnsAttributeCount = spreadCheetahAttributes.RemoveAll(x => x.TryGetInheritColumnsAttribute(out _));

        return inheritColumnsAttributeCount > 0 && type.BaseType is { } baseType
            ? [.. GetAllAttributes(baseType), .. spreadCheetahAttributes]
            : spreadCheetahAttributes;
    }

    private bool TryGetDefaultColumnWidthAttribute(
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

        Result.DefaultColumnWidth = new ColumnWidth(attributeValue);
        return true;
    }
}
