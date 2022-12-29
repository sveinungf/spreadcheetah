using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal static class SymbolExtensions
{
    public static string ToTypeString(this INamedTypeSymbol symbol)
    {
        var result = symbol.ToString();
        return symbol.IsReferenceType ? result + '?' : result;
    }
}
