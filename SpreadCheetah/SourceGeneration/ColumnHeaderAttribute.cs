namespace SpreadCheetah.SourceGeneration;

[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ColumnHeaderAttribute(string name) : Attribute;
