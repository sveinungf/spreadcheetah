namespace SpreadCheetah.SourceGenerator.Models;

internal readonly record struct ColumnHeader
{
    public string? RawString { get; }
    public ColumnHeaderPropertyReference? PropertyReference { get; }

    public ColumnHeader(string rawString) => RawString = rawString;
    public ColumnHeader(ColumnHeaderPropertyReference propertyReference) => PropertyReference = propertyReference;
}