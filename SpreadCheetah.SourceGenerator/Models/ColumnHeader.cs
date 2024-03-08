namespace SpreadCheetah.SourceGenerator.Models;

internal readonly record struct ColumnHeader
{
    public string? RawString { get; }
    public (Type Type, string PropertyName)? TypeProperty { get; }

    public ColumnHeader(string rawString) => RawString = rawString;
    public ColumnHeader(Type type, string propertyName) => TypeProperty = (type, propertyName);
}