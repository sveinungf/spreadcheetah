using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed class DiagnosticsReporter(SymbolAnalysisContext context) : IDiagnosticsReporter
{
    public void ReportInvalidArgument(AttributeData attribute, CancellationToken token)
    {
        if (attribute.ApplicationSyntaxReference?.GetSyntax(token) is not AttributeSyntax attributeSyntax)
            return;
        if (attributeSyntax.ArgumentList?.Arguments.FirstOrDefault() is not { } arg)
            return;
        if (attribute.AttributeClass?.Name is not { } name)
            return;

        var diagnostic = Diagnostics.InvalidAttributeArgument(arg.GetLocation(), name);
        context.ReportDiagnostic(diagnostic);
    }

    // TODO: Set correct location
    public void ReportInvalidPropertyReference(AttributeData attribute, string propertyName, string typeFullName, CancellationToken token)
    {
        context.ReportDiagnostic(Diagnostics.InvalidColumnHeaderPropertyReference(attribute, propertyName, typeFullName, token));
    }

    // TODO: Set correct location
    public void ReportUnsupportedType(AttributeData attribute, ITypeSymbol propertyType, CancellationToken token)
    {
        var typeFullName = propertyType.ToDisplayString();
        context.ReportDiagnostic(Diagnostics.UnsupportedTypeForAttribute(attribute, typeFullName, token));
    }

    // TODO: Set correct location
    public void ReportTypeMustHaveDefaultConstructor(IPropertySymbol symbol, AttributeData attribute, string typeName, CancellationToken token)
    {
        context.ReportDiagnostic(Diagnostics.AttributeTypeArgumentMustHaveDefaultConstructor(symbol, attribute, typeName, token));
    }

    // TODO: Set correct location
    public void ReportTypeMustInherit(AttributeData attribute, string typeName, string baseClassName, CancellationToken token)
    {
        context.ReportDiagnostic(Diagnostics.AttributeTypeArgumentMustInherit(attribute, typeName, baseClassName, token));
    }
}