using SpreadCheetah.CellValueWriters;
using SpreadCheetah.CellValueWriters.Characters;
using SpreadCheetah.Styling;
using System.Diagnostics;

namespace SpreadCheetah.CellWriters;

internal sealed class StyledCellWriter : ICellWriter<StyledCell>
{
    public static StyledCellWriter Instance { get; } = new();

    public bool TryWrite(in StyledCell cell, CellWriterState state)
    {
        return cell.StyleId is null
            ? CellValueWriter.GetWriter(cell.DataCell.Type).TryWriteCell(cell.DataCell, state)
            : CellValueWriter.GetWriter(cell.DataCell.Type).TryWriteCell(cell.DataCell, cell.StyleId, state.Buffer);
    }

    public bool TryWrite(in StyledCell cell, StyleId styleId, CellWriterState state)
    {
        var actualStyleId = cell.StyleId ?? styleId;
        return CellValueWriter.GetWriter(cell.DataCell.Type).TryWriteCell(cell.DataCell, actualStyleId, state.Buffer);
    }

    public void WriteStartElement(in StyledCell cell, StyleId? styleId, CellWriterState state)
    {
        var actualStyleId = cell.StyleId ?? styleId;
        Debug.Assert(CellValueWriter.GetWriter(cell.DataCell.Type) is StringCellValueWriterBase);
        var result = StringCellValueWriterBase.WriteStartElement(actualStyleId, state.Buffer);
        Debug.Assert(result);
    }

    public bool TryWriteValue(in StyledCell cell, ref int valueIndex, CellWriterState state)
    {
        return state.Buffer.WriteLongStringXmlEncoded(cell.DataCell.Value.StringOrPrimitive.StringValue, ref valueIndex);
    }

    public bool TryWriteEndElement(in StyledCell cell, CellWriterState state)
    {
        return CellValueWriter.GetWriter(cell.DataCell.Type).TryWriteEndElement(state.Buffer);
    }
}