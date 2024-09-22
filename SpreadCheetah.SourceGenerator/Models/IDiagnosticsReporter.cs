using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Models;

internal interface IDiagnosticsReporter
{
    void ReportInvalidArgument(AttributeData attribute, CancellationToken token);

    void ReportInvalidPropertyReference(AttributeData attribute, string propertyName, string typeFullName, CancellationToken token);

    void ReportUnsupportedType(AttributeData attribute, ITypeSymbol propertyType, CancellationToken token);

    void ReportTypeMustHaveDefaultConstructor(IPropertySymbol symbol, AttributeData attribute, string typeName, CancellationToken token);

    void ReportTypeMustInherit(AttributeData attribute, string typeName, string baseClassName, CancellationToken token);
}
