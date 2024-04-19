using System.Runtime.InteropServices;

namespace SpreadCheetah.CellValues;

[StructLayout(LayoutKind.Explicit)]
internal readonly struct PrimitiveCellValue
{
    [FieldOffset(0)] public readonly int IntValue;
    [FieldOffset(0)] public readonly float FloatValue;
    [FieldOffset(0)] public readonly double DoubleValue;

#if NET5_0_OR_GREATER
    public PrimitiveCellValue(int value)
    {
        System.Runtime.CompilerServices.Unsafe.SkipInit(out this);
        IntValue = value;
    }

    public PrimitiveCellValue(float value)
    {
        System.Runtime.CompilerServices.Unsafe.SkipInit(out this);
        FloatValue = value;
    }

    public PrimitiveCellValue(double value)
    {
        System.Runtime.CompilerServices.Unsafe.SkipInit(out this);
        DoubleValue = value;
    }
#else
    public PrimitiveCellValue(int value) : this() => IntValue = value;
    public PrimitiveCellValue(float value) : this() => FloatValue = value;
    public PrimitiveCellValue(double value) : this() => DoubleValue = value;
#endif
}
