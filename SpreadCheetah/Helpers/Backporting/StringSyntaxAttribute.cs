#if !NET7_0_OR_GREATER
namespace System.Diagnostics.CodeAnalysis;

[AttributeUsage(AttributeTargets.Parameter | AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
internal sealed class StringSyntaxAttribute : Attribute
{
    public StringSyntaxAttribute(string syntax) => Syntax = syntax;

    public string Syntax { get; }

    public const string Regex = nameof(Regex);
}
#endif