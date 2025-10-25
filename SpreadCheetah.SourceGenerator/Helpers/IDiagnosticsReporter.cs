using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal interface IDiagnosticsReporter
{
    void ReportAttributeCombinationNotSupported(AttributeData attribute, string otherAttribute, CancellationToken token);

    void ReportDuplicateColumnOrdering(AttributeData attribute, CancellationToken token);

    void ReportInvalidArgument(AttributeData attribute, CancellationToken token);

    void ReportUnsupportedPropertyReference(AttributeData attribute, string propertyName, string typeFullName, CancellationToken token);

    void ReportUnsupportedPropertyReference(IPropertySymbol property, string referencedPropertyName, string typeFullName);

    void ReportMissingPropertyReference(AttributeData attribute, string propertyName, string typeFullName, CancellationToken token);

    void ReportMissingPropertyReference(IPropertySymbol property, string referencedPropertyName, string typeFullName);

    void ReportNoPropertiesFound(AttributeData attribute, INamedTypeSymbol rowType, CancellationToken token);

    void ReportUnsupportedPropertyType(AttributeData attribute, INamedTypeSymbol rowType, IPropertySymbol property, CancellationToken token);

    void ReportUnsupportedPropertyTypeForAttribute(AttributeData attribute, ITypeSymbol propertyType, CancellationToken token);

    void ReportTypeMustHaveDefaultConstructor(AttributeData attribute, string typeName, CancellationToken token);

    void ReportTypeMustInherit(AttributeData attribute, string typeName, string baseClassName, CancellationToken token);
}
