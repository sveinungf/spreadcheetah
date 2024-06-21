namespace SpreadCheetah.CellValues;

internal readonly struct StringOrPrimitiveCellValue
{
    public readonly string? StringValue;
    public readonly PrimitiveCellValue PrimitiveValue;

#if NET5_0_OR_GREATER
    public StringOrPrimitiveCellValue()
    {
        System.Runtime.CompilerServices.Unsafe.SkipInit(out this);
    }
#endif

    public StringOrPrimitiveCellValue(string? value) : this() => StringValue = value;
    public StringOrPrimitiveCellValue(PrimitiveCellValue value) : this() => PrimitiveValue = value;
}