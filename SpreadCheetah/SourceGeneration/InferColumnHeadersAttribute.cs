namespace SpreadCheetah.SourceGeneration;

// TODO: Inherited = true or false?
[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
public sealed class InferColumnHeadersAttribute(Type type) : Attribute
{
    public Type? Type { get; } = type;
    public string? Prefix { get; set; }
    public string? Suffix { get; set; }
}
