using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed class DiagnosticsReporter(SymbolAnalysisContext context) : IDiagnosticsReporter
{
    public void ReportInvalidArgument(AttributeData attribute, CancellationToken token)
    {
        if (!TryGetArgument(attribute, token, out var arg) || arg is null)
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

    public void ReportTypeMustHaveDefaultConstructor(AttributeData attribute, string typeName, CancellationToken token)
    {
        if (!TryGetArgument(attribute, token, out var arg) || arg is null)
            return;

        var diagnostic = Diagnostics.AttributeTypeArgumentMustHaveDefaultConstructor(arg.GetLocation(), typeName);
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportTypeMustInherit(AttributeData attribute, string typeName, string baseClassName, CancellationToken token)
    {
        if (!TryGetArgument(attribute, token, out var arg) || arg is null)
            return;
        if (attribute.AttributeClass?.Name is not { } name)
            return;

        var diagnostic = Diagnostics.AttributeTypeArgumentMustInherit(arg.GetLocation(), typeName, name, baseClassName);
        context.ReportDiagnostic(diagnostic);
    }

    private static bool TryGetArgument(AttributeData attribute, CancellationToken token, out AttributeArgumentSyntax? argument)
    {
        argument = null;

        if (attribute.ApplicationSyntaxReference?.GetSyntax(token) is not AttributeSyntax attributeSyntax)
            return false;
        if (attributeSyntax.ArgumentList?.Arguments.FirstOrDefault() is not { } arg)
            return false;

        argument = arg;
        return true;
    }
}