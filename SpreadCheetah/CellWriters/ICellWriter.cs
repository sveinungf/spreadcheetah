using SpreadCheetah.Styling;

namespace SpreadCheetah.CellWriters;

internal interface ICellWriter<T>
    where T : struct
{
    bool TryWrite(in T cell, CellWriterState state);
    bool TryWrite(in T cell, StyleId styleId, CellWriterState state);
    void WriteStartElement(in T cell, StyleId? styleId, CellWriterState state);
    bool TryWriteValue(in T cell, ref int valueIndex, CellWriterState state);
    bool TryWriteEndElement(in T cell, CellWriterState state);
}
