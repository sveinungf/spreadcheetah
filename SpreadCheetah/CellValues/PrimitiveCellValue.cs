using System.Runtime.InteropServices;

namespace SpreadCheetah.CellValues;

[StructLayout(LayoutKind.Explicit)]
internal readonly struct PrimitiveCellValue
{
    [FieldOffset(0)] public readonly int IntValue;
    [FieldOffset(0)] public readonly float FloatValue;
    [FieldOffset(0)] public readonly double DoubleValue;

    public PrimitiveCellValue()
    {
#if NET5_0_OR_GREATER
        System.Runtime.CompilerServices.Unsafe.SkipInit(out this);
#endif
    }

    public PrimitiveCellValue(int value) : this() => IntValue = value;
    public PrimitiveCellValue(float value) : this() => FloatValue = value;
    public PrimitiveCellValue(double value) : this() => DoubleValue = value;
}
