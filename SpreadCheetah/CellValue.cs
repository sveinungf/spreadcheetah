using System.Runtime.InteropServices;

namespace SpreadCheetah
{
    [StructLayout(LayoutKind.Explicit)]
    internal readonly struct CellValue
    {
        [FieldOffset(0)] public readonly int IntValue;
        [FieldOffset(0)] public readonly float FloatValue;
        [FieldOffset(0)] public readonly double DoubleValue;

        public CellValue(int value) : this() => IntValue = value;
        public CellValue(float value) : this() => FloatValue = value;
        public CellValue(double value) : this() => DoubleValue = value;
    }
}
