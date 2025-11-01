using Microsoft.CodeAnalysis;
using SpreadCheetah.SourceGenerator.Helpers;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed class InferColumnHeaderDiagnosticData : IColumnHeaderDiagnosticData
{
    public required IPropertySymbol Property { get; init; }
    public required string ReferencedPropertyName { get; init; }
    public required string TypeFullName { get; init; }

    public Location? GetLocation(CancellationToken token)
    {
        return Property.Locations.FirstOrDefault();
    }
}