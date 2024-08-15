using SpreadCheetah.SourceGenerator.Helpers;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowType(
    string Name,
    string FullName,
    bool IsReferenceType,
    bool HasStyleAttribute,
    LocationInfo? WorksheetRowAttributeLocation,
    EquatableArray<RowTypeProperty> Properties,
    EquatableArray<string> UnsupportedPropertyTypeNames,
    EquatableArray<DiagnosticInfo> DiagnosticInfos)
{
    public string CellType { get; } = HasStyleAttribute ? "StyledCell" : "DataCell";
    public string FullNameWithNullableAnnotation => IsReferenceType ? $"{FullName}?" : FullName;
}