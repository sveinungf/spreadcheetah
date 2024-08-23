namespace SpreadCheetah.SourceGenerator.Models;

internal readonly record struct CellValueMapper(
    string CellValueMapperTypeName,
    string GenericName,
    LocationInfo LocationInfo);