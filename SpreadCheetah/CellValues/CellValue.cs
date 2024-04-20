using System.Runtime.InteropServices;

namespace SpreadCheetah.CellValues;

[StructLayout(LayoutKind.Explicit)]
internal readonly struct CellValue
{
    [FieldOffset(0)] public readonly StringOrPrimitiveCellValue StringOrPrimitive;
    [FieldOffset(0)] public readonly ReadOnlyMemory<char> Memory;

    public CellValue()
    {
#if NET5_0_OR_GREATER
        System.Runtime.CompilerServices.Unsafe.SkipInit(out this);
#endif
    }

    public CellValue(StringOrPrimitiveCellValue value) : this() => StringOrPrimitive = value;
    public CellValue(ReadOnlyMemory<char> value) : this() => Memory = value;
}