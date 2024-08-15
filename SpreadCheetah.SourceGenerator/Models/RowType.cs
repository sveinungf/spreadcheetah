using SpreadCheetah.SourceGenerator.Helpers;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowType(
    string Name,
    string FullName,
    bool IsReferenceType,
    int PropertiesWithStyleAttributes,
    LocationInfo? WorksheetRowAttributeLocation,
    EquatableArray<RowTypeProperty> Properties,
    EquatableArray<string> UnsupportedPropertyTypeNames,
    EquatableArray<DiagnosticInfo> DiagnosticInfos)
{
    public string CellType { get; } = PropertiesWithStyleAttributes > 0 ? "StyledCell" : "DataCell";
    public string FullNameWithNullableAnnotation => IsReferenceType ? $"{FullName}?" : FullName;
}