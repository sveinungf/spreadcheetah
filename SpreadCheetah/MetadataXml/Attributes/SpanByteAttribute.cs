using System.Runtime.InteropServices;

namespace SpreadCheetah.MetadataXml.Attributes;

[StructLayout(LayoutKind.Auto)]
internal readonly ref struct SpanByteAttribute(ReadOnlySpan<byte> attributeName, ReadOnlySpan<byte> value)
{
    public readonly ReadOnlySpan<byte> AttributeName { get; } = attributeName;

    public readonly ReadOnlySpan<byte> Value { get; } = value;
}