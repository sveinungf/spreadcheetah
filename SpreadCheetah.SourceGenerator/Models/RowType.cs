using SpreadCheetah.SourceGenerator.Helpers;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowType(
    string Name,
    string FullName,
    bool IsReferenceType,
    CellType CellType,
    bool HasStyle,
    ColumnWidth? DefaultColumnWidth,
    EquatableArray<RowTypeProperty> Properties)
{
    public string FullNameWithNullableAnnotation => IsReferenceType ? $"{FullName}?" : FullName;
}