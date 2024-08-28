namespace SpreadCheetah.SourceGenerator.Models;

internal readonly record struct CellValueConverter(
    string CellValueConverterTypeName,
    string GenericName,
    LocationInfo LocationInfo);