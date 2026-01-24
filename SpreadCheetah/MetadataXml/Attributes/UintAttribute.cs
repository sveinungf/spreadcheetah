using System.Runtime.InteropServices;

namespace SpreadCheetah.MetadataXml.Attributes;

[StructLayout(LayoutKind.Auto)]
internal readonly ref struct UintAttribute(ReadOnlySpan<byte> attributeName, uint? value)
{
    public readonly ReadOnlySpan<byte> AttributeName { get; } = attributeName;

    public readonly uint? Value { get; } = value;
}