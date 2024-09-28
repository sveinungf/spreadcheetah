using SpreadCheetah.SourceGenerator.Helpers;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowType(
    string Name,
    string FullName,
    bool IsReferenceType,
    int PropertiesWithStyleAttributes,
    EquatableArray<RowTypeProperty> Properties)
{
    public string CellType { get; } = PropertiesWithStyleAttributes > 0 ? "StyledCell" : "DataCell";
    public string FullNameWithNullableAnnotation => IsReferenceType ? $"{FullName}?" : FullName;
}