using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator
{
    internal static class Diagnostics
    {
        public static readonly DiagnosticDescriptor NoPropertiesFoundWarning = new DiagnosticDescriptor(
            id: "SPCH101",
            title: "Missing properties with public getters",
            messageFormat: "The type '{0}' has no properties with public getters. This will cause an empty row to be added.",
            category: "SpreadCheetah.SourceGenerator",
            DiagnosticSeverity.Warning,
            isEnabledByDefault: true);
    }
}
