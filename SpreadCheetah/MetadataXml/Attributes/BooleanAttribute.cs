using System.Runtime.InteropServices;

namespace SpreadCheetah.MetadataXml.Attributes;

[StructLayout(LayoutKind.Auto)]
internal readonly ref struct BooleanAttribute(ReadOnlySpan<byte> attributeName, bool? value)
{
    public readonly ReadOnlySpan<byte> AttributeName { get; } = attributeName;

    public readonly bool? Value { get; } = value;
}