using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal interface IDiagnosticsReporter
{
    void ReportInvalidArgument(AttributeData attribute, CancellationToken token);

    void ReportInvalidPropertyReference(AttributeData attribute, string propertyName, string typeFullName, CancellationToken token);

    void ReportUnsupportedPropertyTypeForAttribute(AttributeData attribute, ITypeSymbol propertyType, CancellationToken token);

    void ReportTypeMustHaveDefaultConstructor(AttributeData attribute, bool genericAttribute, string typeName, CancellationToken token);

    void ReportTypeMustInherit(AttributeData attribute, bool genericAttribute, string typeName, string baseClassName, CancellationToken token);
}
