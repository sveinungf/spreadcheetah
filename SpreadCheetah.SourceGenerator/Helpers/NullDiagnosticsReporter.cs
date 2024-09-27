using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Helpers;

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

    public void ReportTypeMustHaveDefaultConstructor(AttributeData attribute, string typeName, CancellationToken token)
    {
    }

    public void ReportTypeMustInherit(AttributeData attribute, string typeName, string baseClassName, CancellationToken token)
    {
    }

    public void ReportUnsupportedType(AttributeData attribute, ITypeSymbol propertyType, CancellationToken token)
    {
    }
}
