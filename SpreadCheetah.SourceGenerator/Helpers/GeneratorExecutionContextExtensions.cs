using Microsoft.CodeAnalysis;
using System.Collections.Generic;

namespace SpreadCheetah.SourceGenerator.Helpers
{
    internal static class GeneratorExecutionContextExtensions
    {
        public static void ReportDiagnostics(
            this GeneratorExecutionContext context,
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
}
