using SpreadCheetah.CellValueWriters;
using SpreadCheetah.CellValueWriters.Characters;
using SpreadCheetah.Styling;
using System.Diagnostics;

namespace SpreadCheetah.CellWriters;

internal sealed class DataCellWriter : ICellWriter<DataCell>
{
    public static DataCellWriter Instance { get; } = new();

    public bool TryWrite(in DataCell cell, CellWriterState state)
    {
        return CellValueWriter.GetWriter(cell.Type).TryWriteCell(cell, state);
    }

    public bool TryWrite(in DataCell cell, StyleId styleId, CellWriterState state)
    {
        return CellValueWriter.GetWriter(cell.Type).TryWriteCell(cell, styleId, state.Buffer);
    }

    public void WriteStartElement(in DataCell cell, StyleId? styleId, CellWriterState state)
    {
        Debug.Assert(CellValueWriter.GetWriter(cell.Type) is StringCellValueWriterBase);
        var result = StringCellValueWriterBase.WriteStartElement(styleId, state.Buffer);
        Debug.Assert(result);
    }

    public bool TryWriteValue(in DataCell cell, ref int valueIndex, CellWriterState state)
    {
        return state.Buffer.WriteLongString(cell.Value.StringOrPrimitive.StringValue, ref valueIndex);
    }

    public bool TryWriteEndElement(in DataCell cell, CellWriterState state)
    {
        return CellValueWriter.GetWriter(cell.Type).TryWriteEndElement(state.Buffer);
    }
}
