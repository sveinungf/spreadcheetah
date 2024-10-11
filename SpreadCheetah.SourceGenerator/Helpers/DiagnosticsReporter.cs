using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed class DiagnosticsReporter(SymbolAnalysisContext context) : IDiagnosticsReporter
{
    public void ReportAttributeCombinationNotSupported(AttributeData attribute, string otherAttribute, CancellationToken token)
    {
        var attributeSyntax = attribute.ApplicationSyntaxReference?.GetSyntax(token);
        var name = attribute.AttributeClass?.Name;
        var diagnostic = Diagnostics.AttributeCombinationNotSupported(attributeSyntax?.GetLocation(), name, otherAttribute);
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportDuplicateColumnOrdering(AttributeData attribute, CancellationToken token)
    {
        var attributeSyntax = attribute.ApplicationSyntaxReference?.GetSyntax(token);
        var diagnostic = Diagnostics.DuplicateColumnOrder(attributeSyntax?.GetLocation());
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportInvalidArgument(AttributeData attribute, CancellationToken token)
    {
        var location = GetArgument(attribute, token)?.GetLocation();
        var name = attribute.AttributeClass?.Name;
        var diagnostic = Diagnostics.InvalidAttributeArgument(location, name);
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportInvalidPropertyReference(AttributeData attribute, string propertyName, string typeFullName, CancellationToken token)
    {
        var argList = GetAttributeSyntax(attribute, token)?.ArgumentList;
        var diagnostic = Diagnostics.InvalidColumnHeaderPropertyReference(argList?.GetLocation(), propertyName, typeFullName);
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportNoPropertiesFound(AttributeData attribute, INamedTypeSymbol rowType, CancellationToken token)
    {
        var location = GetArgument(attribute, token)?.GetLocation();
        var diagnostic = Diagnostics.NoPropertiesFound(location, rowType.Name);
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportUnsupportedPropertyType(AttributeData attribute, bool genericAttribute, INamedTypeSymbol rowType, IPropertySymbol property, CancellationToken token)
    {
        var location = genericAttribute
            ? GetTypeArgument(attribute, token)?.GetLocation()
            : GetArgument(attribute, token)?.GetLocation();

        var diagnostic = Diagnostics.UnsupportedTypeForCellValue(location, rowType.Name, property.Name, property.Type.Name);
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportUnsupportedPropertyTypeForAttribute(AttributeData attribute, ITypeSymbol propertyType, CancellationToken token)
    {
        var syntaxNode = attribute.ApplicationSyntaxReference?.GetSyntax(token);
        var name = attribute.AttributeClass?.Name;
        var diagnostic = Diagnostics.UnsupportedTypeForAttribute(syntaxNode?.GetLocation(), name, propertyType.ToDisplayString());
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportTypeMustHaveDefaultConstructor(AttributeData attribute, bool genericAttribute, string typeName, CancellationToken token)
    {
        var location = genericAttribute
            ? GetTypeArgument(attribute, token)?.GetLocation()
            : GetArgument(attribute, token)?.GetLocation();

        var diagnostic = Diagnostics.AttributeTypeArgumentMustHaveDefaultConstructor(location, typeName);
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportTypeMustInherit(AttributeData attribute, bool genericAttribute, string typeName, string baseClassName, CancellationToken token)
    {
        var location = genericAttribute
            ? GetTypeArgument(attribute, token)?.GetLocation()
            : GetArgument(attribute, token)?.GetLocation();

        var name = attribute.AttributeClass?.Name;
        var diagnostic = Diagnostics.AttributeTypeArgumentMustInherit(location, typeName, name, baseClassName);
        context.ReportDiagnostic(diagnostic);
    }

    private static AttributeSyntax? GetAttributeSyntax(AttributeData attribute, CancellationToken token)
    {
        var syntaxNode = attribute.ApplicationSyntaxReference?.GetSyntax(token);
        return syntaxNode as AttributeSyntax;
    }

    private static AttributeArgumentSyntax? GetArgument(AttributeData attribute, CancellationToken token)
    {
        var argList = GetAttributeSyntax(attribute, token)?.ArgumentList;
        return argList?.Arguments.FirstOrDefault();
    }

    private static TypeSyntax? GetTypeArgument(AttributeData attribute, CancellationToken token)
    {
        var attributeSyntax = GetAttributeSyntax(attribute, token);
        var name = attributeSyntax?.Name as GenericNameSyntax;
        return name?.TypeArgumentList.Arguments.FirstOrDefault();
    }
}