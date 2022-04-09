using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed class ContextClass
{
    public ITypeSymbol ContextClassType { get; }
    public ISet<INamedTypeSymbol> RowTypes { get; }

    public ContextClass(ITypeSymbol contextClassType, ISet<INamedTypeSymbol> rowTypes)
    {
        ContextClassType = contextClassType;
        RowTypes = rowTypes;
    }
}
