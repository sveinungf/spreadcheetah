using System.Runtime.InteropServices;

namespace SpreadCheetah.MetadataXml.Attributes;

[StructLayout(LayoutKind.Auto)]
internal readonly ref struct DoubleAttribute(ReadOnlySpan<byte> attributeName, double? value)
{
    public readonly ReadOnlySpan<byte> AttributeName { get; } = attributeName;

    public readonly double? Value { get; } = value;
}
