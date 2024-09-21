using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Models;

internal interface IDiagnosticsReporter
{
    void ReportInvalidArgument(AttributeData attribute, int value, CancellationToken token);

    void ReportInvalidArgument(AttributeData attribute, double value, CancellationToken token);

    void ReportInvalidArgument(AttributeData attribute, string value, CancellationToken token);

    void ReportInvalidPropertyReference(AttributeData attribute, string propertyName, string typeFullName, CancellationToken token);

    void ReportUnsupportedType(AttributeData attribute, ITypeSymbol propertyType, CancellationToken token);

    void ReportTypeMustHaveDefaultConstructor(AttributeData attribute, string typeName, CancellationToken token);

    void ReportTypeMustInherit(AttributeData attribute, string typeName, string baseClassName, CancellationToken token);
}
