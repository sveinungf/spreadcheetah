using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowProperty(
    IPropertySymbol PropertySymbol,
    InferColumnHeadersInfo? InferColumnHeadersInfo);