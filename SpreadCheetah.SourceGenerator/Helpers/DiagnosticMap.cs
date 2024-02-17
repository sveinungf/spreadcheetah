using Microsoft.CodeAnalysis;
using SpreadCheetah.SourceGenerator.Models;

namespace SpreadCheetah.SourceGenerator.Helpers;

internal static class DiagnosticMap
{
    public static Diagnostic ToDiagnostic(this DiagnosticInfo info)
    {
        return Diagnostic.Create(info.Descriptor, info.Location?.ToLocation(), info.MessageArgs.GetArray());
    }
}
