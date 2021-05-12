using System.Runtime.InteropServices;

namespace SpreadCheetah
{
    [StructLayout(LayoutKind.Explicit)]
    internal readonly struct CellValue
    {
        [FieldOffset(0)] public readonly int IntValue;
        [FieldOffset(0)] public readonly float FloatValue;
        [FieldOffset(0)] public readonly double DoubleValue;

#if NET5_0_OR_GREATER
        public CellValue(int value)
        {
            System.Runtime.CompilerServices.Unsafe.SkipInit(out this);
            IntValue = value;
        }

        public CellValue(float value)
        {
            System.Runtime.CompilerServices.Unsafe.SkipInit(out this);
            FloatValue = value;
        }

        public CellValue(double value)
        {
            System.Runtime.CompilerServices.Unsafe.SkipInit(out this);
            DoubleValue = value;
        }
#else
        public CellValue(int value) : this() => IntValue = value;
        public CellValue(float value) : this() => FloatValue = value;
        public CellValue(double value) : this() => DoubleValue = value;
#endif
    }
}
