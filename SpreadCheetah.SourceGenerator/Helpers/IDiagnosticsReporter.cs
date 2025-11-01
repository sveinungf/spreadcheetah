using Microsoft.CodeAnalysis;
using SpreadCheetah.SourceGenerator.Models;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal interface IDiagnosticsReporter
{
    void ReportAttributeCombinationNotSupported(AttributeData attribute, string otherAttribute, CancellationToken token);

    void ReportDuplicateColumnOrdering(AttributeData attribute, CancellationToken token);

    void ReportInvalidArgument(AttributeData attribute, CancellationToken token);

    void ReportMissingPropertyForColumnHeader(IColumnHeaderDiagnosticData data, CancellationToken token);

    void ReportNonPublicPropertyForColumnHeader(IColumnHeaderDiagnosticData data, CancellationToken token);

    void ReportUnsupportedPropertyForColumnHeader(IColumnHeaderDiagnosticData data, CancellationToken token);

    void ReportNoPropertiesFound(AttributeData attribute, INamedTypeSymbol rowType, CancellationToken token);

    void ReportUnsupportedPropertyType(AttributeData attribute, INamedTypeSymbol rowType, IPropertySymbol property, CancellationToken token);

    void ReportUnsupportedPropertyTypeForAttribute(AttributeData attribute, ITypeSymbol propertyType, CancellationToken token);

    void ReportTypeMustHaveDefaultConstructor(AttributeData attribute, string typeName, CancellationToken token);

    void ReportTypeMustInherit(AttributeData attribute, string typeName, string baseClassName, CancellationToken token);
}
