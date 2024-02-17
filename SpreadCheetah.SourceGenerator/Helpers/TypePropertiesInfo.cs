using SpreadCheetah.SourceGenerator.Models;

namespace SpreadCheetah.SourceGenerator.Helpers;
internal sealed record TypePropertiesInfo(
    EquatableArray<RowTypeProperty> Properties,
    EquatableArray<string> UnsupportedPropertyTypeNames,
    EquatableArray<DiagnosticInfo> DiagnosticInfos);