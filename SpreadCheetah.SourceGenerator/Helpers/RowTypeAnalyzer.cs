using Microsoft.CodeAnalysis;
using SpreadCheetah.SourceGenerator.Extensions;
using SpreadCheetah.SourceGenerator.Models;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed class RowTypeAnalyzer(
    IDiagnosticsReporter diagnostics,
    ITypeSymbol rowType)
{
    public RowTypeAttributeData Result { get; } = new();

    public void Analyze(CancellationToken token)
    {
        var attributes = rowType
            .GetAttributes()
            .Where(x => x.AttributeClass is { } c && c.HasSpreadCheetahSrcGenNamespace());

        foreach (var attribute in attributes)
        {
            _ = attribute.AttributeClass?.MetadataName switch
            {
                Attributes.DefaultColumnWidth => TryGetDefaultColumnWidthAttribute(attribute, token),
                _ => false
            };
        }
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
