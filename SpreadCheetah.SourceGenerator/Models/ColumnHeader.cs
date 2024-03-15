namespace SpreadCheetah.SourceGenerator.Models;

internal readonly record struct ColumnHeader
{
    public string? RawString { get; }
    public (string TypeFullName, string PropertyName)? TypeProperty { get; }

    public ColumnHeader(string rawString) => RawString = rawString;
    public ColumnHeader(string typeFullName, string propertyName) => TypeProperty = (typeFullName, propertyName);
}