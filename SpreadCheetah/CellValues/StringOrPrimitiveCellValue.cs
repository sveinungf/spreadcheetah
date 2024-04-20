namespace SpreadCheetah.CellValues;

internal readonly struct StringOrPrimitiveCellValue
{
    public readonly string? StringValue;
    public readonly PrimitiveCellValue PrimitiveValue;

    public StringOrPrimitiveCellValue()
    {
#if NET5_0_OR_GREATER
        System.Runtime.CompilerServices.Unsafe.SkipInit(out this);
#endif
    }

    public StringOrPrimitiveCellValue(string? value) : this() => StringValue = value;
    public StringOrPrimitiveCellValue(PrimitiveCellValue value) : this() => PrimitiveValue = value;
}