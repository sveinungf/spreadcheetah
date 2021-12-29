using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal static class SourceProductionContextExtensions
{
    public static void ReportDiagnostics(
        this SourceProductionContext context,
        DiagnosticDescriptor descriptor,
        IEnumerable<Location> locations,
        params object?[]? messageArgs)
    {
        foreach (var location in locations)
        {
            context.ReportDiagnostic(Diagnostic.Create(descriptor, location, messageArgs));
        }
    }
}
