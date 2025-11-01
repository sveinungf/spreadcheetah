using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Models;

internal interface IColumnHeaderDiagnosticData
{
    string ReferencedPropertyName { get; }
    string TypeFullName { get; }
    Location? GetLocation(CancellationToken token);
}