using SpreadCheetah.CellValueWriters;
using SpreadCheetah.CellValueWriters.Characters;
using SpreadCheetah.Styling;
using System.Diagnostics;

namespace SpreadCheetah.CellWriters;

internal sealed class DataCellWithReferenceWriter : ICellWriter<DataCell>
{
    public static DataCellWithReferenceWriter Instance { get; } = new();

    public bool TryWrite(in DataCell cell, CellWriterState state)
    {
        return CellValueWriter.GetWriter(cell.Type).TryWriteCellWithReference(cell, state.DefaultStyling, state);
    }

    public bool TryWrite(in DataCell cell, StyleId styleId, CellWriterState state)
    {
        return CellValueWriter.GetWriter(cell.Type).TryWriteCellWithReference(cell, styleId, state);
    }

    public void WriteStartElement(in DataCell cell, StyleId? styleId, CellWriterState state)
    {
        Debug.Assert(CellValueWriter.GetWriter(cell.Type) is StringCellValueWriterBase);
        var result = StringCellValueWriterBase.WriteStartElementWithReference(styleId, state);
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