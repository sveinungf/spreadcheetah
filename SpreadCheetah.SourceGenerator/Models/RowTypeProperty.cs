using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowTypeProperty(
    string Name,
    TypedConstant? ColumnHeaderAttributeValue);