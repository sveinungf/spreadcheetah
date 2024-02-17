using SpreadCheetah.SourceGenerator.Helpers;

namespace SpreadCheetah.SourceGenerator.Models;

internal sealed record RowType(
    string Name,
    string FullName,
    string FullNameWithNullableAnnotation,
    bool IsReferenceType,
    LocationInfo WorksheetRowAttributeLocation,
    EquatableArray<RowTypeProperty> Properties,
    EquatableArray<string> UnsupportedPropertyTypeNames,
    EquatableArray<DiagnosticInfo> DiagnosticInfos);