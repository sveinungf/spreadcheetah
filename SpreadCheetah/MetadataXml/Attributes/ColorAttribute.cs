using System.Drawing;
using System.Runtime.InteropServices;

namespace SpreadCheetah.MetadataXml.Attributes;

[StructLayout(LayoutKind.Auto)]
internal readonly ref struct ColorAttribute(ReadOnlySpan<byte> attributeName, Color? value)
{
    public readonly ReadOnlySpan<byte> AttributeName { get; } = attributeName;

    public readonly Color? Value { get; } = value;
}