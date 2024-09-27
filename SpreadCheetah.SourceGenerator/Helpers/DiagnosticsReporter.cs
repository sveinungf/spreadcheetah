using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed class DiagnosticsReporter(SymbolAnalysisContext context) : IDiagnosticsReporter
{
    public void ReportAttributeCombinationNotSupported(AttributeData attribute, string otherAttribute, CancellationToken token)
    {
        if (attribute.ApplicationSyntaxReference?.GetSyntax(token) is not AttributeSyntax attributeSyntax)
            return;
        if (attribute.AttributeClass?.Name is not { } name)
            return;

        var diagnostic = Diagnostics.AttributeCombinationNotSupported(attributeSyntax.GetLocation(), name, otherAttribute);
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportDuplicateColumnOrdering(AttributeData attribute, CancellationToken token)
    {
        if (attribute.ApplicationSyntaxReference?.GetSyntax(token) is not AttributeSyntax attributeSyntax)
            return;

        var diagnostic = Diagnostics.DuplicateColumnOrder(attributeSyntax.GetLocation());
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportInvalidArgument(AttributeData attribute, CancellationToken token)
    {
        if (!TryGetArgument(attribute, token, out var arg) || arg is null)
            return;
        if (attribute.AttributeClass?.Name is not { } name)
            return;

        var diagnostic = Diagnostics.InvalidAttributeArgument(arg.GetLocation(), name);
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportInvalidPropertyReference(AttributeData attribute, string propertyName, string typeFullName, CancellationToken token)
    {
        if (!TryGetArgumentList(attribute, token, out var argList) || argList is null)
            return;

        var diagnostic = Diagnostics.InvalidColumnHeaderPropertyReference(argList.GetLocation(), propertyName, typeFullName);
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportUnsupportedType(AttributeData attribute, ITypeSymbol propertyType, CancellationToken token)
    {
        if (attribute.ApplicationSyntaxReference?.GetSyntax(token) is not AttributeSyntax attributeSyntax)
            return;
        if (attribute.AttributeClass?.Name is not { } name)
            return;

        var diagnostic = Diagnostics.UnsupportedTypeForAttribute(attributeSyntax.GetLocation(), name, propertyType.ToDisplayString());
        context.ReportDiagnostic(diagnostic);
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

    private static bool TryGetArgumentList(AttributeData attribute, CancellationToken token, out AttributeArgumentListSyntax? argumentList)
    {
        argumentList = null;

        if (attribute.ApplicationSyntaxReference?.GetSyntax(token) is not AttributeSyntax attributeSyntax)
            return false;
        if (attributeSyntax.ArgumentList is not { } argList)
            return false;

        argumentList = argList;
        return true;
    }

    private static bool TryGetArgument(AttributeData attribute, CancellationToken token, out AttributeArgumentSyntax? argument)
    {
        argument = null;

        if (!TryGetArgumentList(attribute, token, out var argList))
            return false;
        if (argList?.Arguments.FirstOrDefault() is not { } arg)
            return false;

        argument = arg;
        return true;
    }
}