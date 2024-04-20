namespace SpreadCheetah.CellValueWriters;

internal enum CellWriterType : byte
{
    Null,
    String = 1,
    Integer = 2,
    Float = 3,
    Double = 4,
    NullDateTime = 5,
    DateTime = 6,
    FalseBoolean = 7,
    TrueBoolean = 8
}