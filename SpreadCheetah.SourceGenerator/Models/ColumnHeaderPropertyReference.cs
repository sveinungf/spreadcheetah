namespace SpreadCheetah.SourceGenerator.Models;

internal readonly record struct ColumnHeaderPropertyReference(
    string TypeFullName,
    string PropertyName);