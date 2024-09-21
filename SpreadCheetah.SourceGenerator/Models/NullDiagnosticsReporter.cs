using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed class NullDiagnosticsReporter : IDiagnosticsReporter
{
    public static readonly NullDiagnosticsReporter Instance = new();

    public void ReportInvalidArgument(AttributeData attribute, int value, CancellationToken token)
    {
    }

    public void ReportInvalidArgument(AttributeData attribute, double value, CancellationToken token)
    {
    }

    public void ReportInvalidArgument(AttributeData attribute, string value, CancellationToken token)
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
