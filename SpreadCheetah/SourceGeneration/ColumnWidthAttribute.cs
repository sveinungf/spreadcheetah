namespace SpreadCheetah.SourceGeneration;

[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ColumnWidthAttribute(int width) : Attribute;
