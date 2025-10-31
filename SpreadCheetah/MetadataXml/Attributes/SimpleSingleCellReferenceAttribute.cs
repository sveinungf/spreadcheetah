using SpreadCheetah.CellReferences;
using System.Runtime.InteropServices;

namespace SpreadCheetah.MetadataXml.Attributes;

[StructLayout(LayoutKind.Auto)]
internal readonly ref struct SimpleSingleCellReferenceAttribute(ReadOnlySpan<byte> attributeName, SimpleSingleCellReference? value)
{
    public readonly ReadOnlySpan<byte> AttributeName { get; } = attributeName;

    public readonly SimpleSingleCellReference? Value { get; } = value;
}