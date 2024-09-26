using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal interface IDiagnosticsReporter
{
    void ReportAttributeCombinationNotSupported(AttributeData attribute, string otherAttribute, CancellationToken token);

    void ReportInvalidArgument(AttributeData attribute, CancellationToken token);

    void ReportInvalidPropertyReference(AttributeData attribute, string propertyName, string typeFullName, CancellationToken token);

    void ReportUnsupportedType(AttributeData attribute, ITypeSymbol propertyType, CancellationToken token);

    void ReportTypeMustHaveDefaultConstructor(AttributeData attribute, string typeName, CancellationToken token);

    void ReportTypeMustInherit(AttributeData attribute, string typeName, string baseClassName, CancellationToken token);
}
