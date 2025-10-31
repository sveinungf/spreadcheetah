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

    public void ReportUnsupportedPropertyForColumnHeader(AttributeData attribute, string propertyName, string typeFullName, CancellationToken token)
    {
        var location = GetLocation(attribute, token);
        var diagnostic = Diagnostics.UnsupportedPropertyForColumnHeader(location, propertyName, typeFullName);
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportUnsupportedPropertyForColumnHeader(IPropertySymbol property, string referencedPropertyName, string typeFullName)
    {
        var location = property.Locations.FirstOrDefault();
        var diagnostic = Diagnostics.UnsupportedPropertyForColumnHeader(location, referencedPropertyName, typeFullName);
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportMissingPropertyForColumnHeader(AttributeData attribute, string propertyName, string typeFullName, CancellationToken token)
    {
        var location = GetLocation(attribute, token);
        var diagnostic = Diagnostics.MissingPropertyForColumnHeader(location, propertyName, typeFullName);
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportMissingPropertyForColumnHeader(IPropertySymbol property, string referencedPropertyName, string typeFullName)
    {
        var location = property.Locations.FirstOrDefault();
        var diagnostic = Diagnostics.MissingPropertyForColumnHeader(location, referencedPropertyName, typeFullName);
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportNoPropertiesFound(AttributeData attribute, INamedTypeSymbol rowType, CancellationToken token)
    {
        var location = GetArgument(attribute, token)?.GetLocation();
        var diagnostic = Diagnostics.NoPropertiesFound(location, rowType.Name);
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportUnsupportedPropertyType(AttributeData attribute, INamedTypeSymbol rowType, IPropertySymbol property, CancellationToken token)
    {
        var location = GetArgument(attribute, token)?.GetLocation();
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

    public void ReportTypeMustHaveDefaultConstructor(AttributeData attribute, string typeName, CancellationToken token)
    {
        var location = GetArgument(attribute, token)?.GetLocation();
        var diagnostic = Diagnostics.AttributeTypeArgumentMustHaveDefaultConstructor(location, typeName);
        context.ReportDiagnostic(diagnostic);
    }

    public void ReportTypeMustInherit(AttributeData attribute, string typeName, string baseClassName, CancellationToken token)
    {
        var location = GetArgument(attribute, token)?.GetLocation();
        var name = attribute.AttributeClass?.Name;
        var diagnostic = Diagnostics.AttributeTypeArgumentMustInherit(location, typeName, name, baseClassName);
        context.ReportDiagnostic(diagnostic);
    }

    private static AttributeArgumentListSyntax? GetArgumentList(AttributeData attribute, CancellationToken token)
    {
        var syntaxNode = attribute.ApplicationSyntaxReference?.GetSyntax(token);
        var attributeSyntax = syntaxNode as AttributeSyntax;
        return attributeSyntax?.ArgumentList;
    }

    private static AttributeArgumentSyntax? GetArgument(AttributeData attribute, CancellationToken token)
    {
        var argList = GetArgumentList(attribute, token);
        return argList?.Arguments.FirstOrDefault();
    }

    private static Location? GetLocation(AttributeData attribute, CancellationToken token)
    {
        var argList = GetArgumentList(attribute, token);
        return argList?.GetLocation();
    }
}