using Microsoft.CodeAnalysis;
using System.Diagnostics.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Helpers;

[ExcludeFromCodeCoverage]
internal sealed class NullDiagnosticsReporter : IDiagnosticsReporter
{
    public static readonly NullDiagnosticsReporter Instance = new();

    public void ReportAttributeCombinationNotSupported(AttributeData attribute, string otherAttribute, CancellationToken token)
    {
    }

    public void ReportDuplicateColumnOrdering(AttributeData attribute, CancellationToken token)
    {
    }

    public void ReportInvalidArgument(AttributeData attribute, CancellationToken token)
    {
    }

    public void ReportInvalidPropertyReference(AttributeData attribute, string propertyName, string typeFullName, CancellationToken token)
    {
    }

    public void ReportNoPropertiesFound(AttributeData attribute, INamedTypeSymbol rowType, CancellationToken token)
    {
    }

    public void ReportTypeMustHaveDefaultConstructor(AttributeData attribute, string typeName, CancellationToken token)
    {
    }

    public void ReportTypeMustInherit(AttributeData attribute, string typeName, string baseClassName, CancellationToken token)
    {
    }

    public void ReportUnsupportedPropertyType(AttributeData attribute, INamedTypeSymbol rowType, IPropertySymbol property, CancellationToken token)
    {
    }

    public void ReportUnsupportedPropertyTypeForAttribute(AttributeData attribute, ITypeSymbol propertyType, CancellationToken token)
    {
    }
}
