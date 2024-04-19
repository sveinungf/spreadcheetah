using System.Runtime.InteropServices;

namespace SpreadCheetah.CellValues;

[StructLayout(LayoutKind.Explicit)]
internal readonly struct CellValue
{
    [FieldOffset(0)] public readonly StringOrPrimitiveCellValue StringOrPrimitive;
    [FieldOffset(0)] public readonly ReadOnlyMemory<char> Memory;

    public CellValue(StringOrPrimitiveCellValue value)
    {
        // TODO: Unsafe.SkipInit?
        StringOrPrimitive = value;
    }

    public CellValue(ReadOnlyMemory<char> value)
    {
        // TODO: Unsafe.SkipInit?
        Memory = value;
    }
}