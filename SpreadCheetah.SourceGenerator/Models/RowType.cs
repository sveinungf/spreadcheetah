using SpreadCheetah.SourceGenerator.Helpers;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowType(
    string Name,
    string FullName,
    bool IsReferenceType,
    bool HasStyleAttributes,
    EquatableArray<RowTypeProperty> Properties)
{
    public string CellType { get; } = HasStyleAttributes ? "StyledCell" : "DataCell";
    public string FullNameWithNullableAnnotation => IsReferenceType ? $"{FullName}?" : FullName;
}