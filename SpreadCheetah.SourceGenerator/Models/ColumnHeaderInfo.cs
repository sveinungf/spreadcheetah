namespace SpreadCheetah.SourceGenerator.Models;

internal readonly record struct ColumnHeaderInfo(
    string? RawString,
    string? FullPropertyReference);