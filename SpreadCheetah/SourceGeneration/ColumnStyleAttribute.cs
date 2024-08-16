namespace SpreadCheetah.SourceGeneration;

[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ColumnStyleAttribute(string styleName) : Attribute; // TODO: Or CellStyle?
