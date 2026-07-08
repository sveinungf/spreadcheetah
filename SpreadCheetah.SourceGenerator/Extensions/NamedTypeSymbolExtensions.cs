using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Extensions;

internal static class NamedTypeSymbolExtensions
{
    extension(INamedTypeSymbol symbol)
    {
        public bool HasSpreadCheetahSrcGenNamespace => symbol is
        {
            ContainingNamespace:
            {
                Name: "SourceGeneration",
                ContainingNamespace:
                {
                    Name: "SpreadCheetah",
                    ContainingNamespace.IsGlobalNamespace: true
                }
            }
        };

        public bool IsWorksheetRowContextBaseClass => symbol is
        {
            IsStatic: false,
            Name: "WorksheetRowContext",
            HasSpreadCheetahSrcGenNamespace: true
        };
    }
}
