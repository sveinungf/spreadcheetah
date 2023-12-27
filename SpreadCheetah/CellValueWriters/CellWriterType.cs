namespace SpreadCheetah.CellValueWriters;

internal enum CellWriterType : byte
{
    Null,
    Integer,
    Float,
    Double,
    DateTime,
    NullDateTime,
    TrueBoolean,
    FalseBoolean,
    String
}