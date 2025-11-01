using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using SpreadCheetah.SourceGenerator.Helpers;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed class ColumnHeaderDiagnosticData : IColumnHeaderDiagnosticData
{
    public required AttributeData Attribute { get; init; }
    public required string ReferencedPropertyName { get; init; }
    public required string TypeFullName { get; init; }

    public Location? GetLocation(CancellationToken token)
    {
        var syntaxNode = Attribute.ApplicationSyntaxReference?.GetSyntax(token);
        var attributeSyntax = syntaxNode as AttributeSyntax;
        var argList = attributeSyntax?.ArgumentList;
        return argList?.GetLocation();
    }
}