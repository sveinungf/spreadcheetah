namespace SpreadCheetah.CellValues;

internal readonly struct StringOrPrimitiveCellValue
{
    public readonly string? StringValue;
    public readonly PrimitiveCellValue PrimitiveValue;

    public StringOrPrimitiveCellValue(string? value)
    {
        // TODO: Unsafe.SkipInit?
        StringValue = value;
    }

    public StringOrPrimitiveCellValue(PrimitiveCellValue value)
    {
        // TODO: Unsafe.SkipInit?
        PrimitiveValue = value;
    }
}