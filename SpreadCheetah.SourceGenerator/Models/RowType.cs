using SpreadCheetah.SourceGenerator.Helpers;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowType
{
    public required string Name { get; init; }
    public required string FullName { get; init; }
    public required bool HasStyle { get; init; }
    public required bool IsReferenceType { get; init; }
    public required CellType CellType { get; init; }
    public required ColumnWidth? DefaultColumnWidth { get; init; }
    public required EquatableArray<RowTypeProperty> Properties { get; init; }

    public string FullNameWithNullableAnnotation => IsReferenceType ? $"{FullName}?" : FullName;
}