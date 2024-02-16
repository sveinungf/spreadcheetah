using Microsoft.CodeAnalysis;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowTypeProperty(
    string Name,
    string TypeName,
    string TypeFullName,
    NullableAnnotation TypeNullableAnnotation,
    SpecialType TypeSpecialType,
    TypedConstant? ColumnHeaderAttributeValue);