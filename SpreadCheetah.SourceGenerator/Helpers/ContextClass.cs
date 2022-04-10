using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal sealed class ContextClass
{
    public ITypeSymbol ContextClassType { get; }
    public Dictionary<INamedTypeSymbol, Location> RowTypes { get; }

    public ContextClass(ITypeSymbol contextClassType, Dictionary<INamedTypeSymbol, Location> rowTypes)
    {
        ContextClassType = contextClassType;
        RowTypes = rowTypes;
    }
}
