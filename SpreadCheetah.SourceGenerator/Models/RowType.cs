namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowType(
    string Name,
    string FullName,
    string FullNameWithNullableAnnotation,
    bool IsReferenceType);