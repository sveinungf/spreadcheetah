using SpreadCheetah.CellValueWriters;
using SpreadCheetah.CellValueWriters.Characters;
using SpreadCheetah.Styling;
using System.Diagnostics;

namespace SpreadCheetah.CellWriters;

internal sealed class StyledCellWithReferenceWriter : ICellWriter<StyledCell>
{
    public static StyledCellWithReferenceWriter Instance { get; } = new();

    public bool TryWrite(in StyledCell cell, CellWriterState state)
    {
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return cell.StyleId is null
            ? writer.TryWriteCellWithReference(cell.DataCell, state)
            : writer.TryWriteCellWithReference(cell.DataCell, cell.StyleId, state);
    }

    public bool TryWrite(in StyledCell cell, StyleId styleId, CellWriterState state)
    {
        var actualStyleId = cell.StyleId ?? styleId;
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return writer.TryWriteCellWithReference(cell.DataCell, actualStyleId, state);
    }

    public void WriteStartElement(in StyledCell cell, StyleId? styleId, CellWriterState state)
    {
        var actualStyleId = cell.StyleId ?? styleId;
        Debug.Assert(CellValueWriter.GetWriter(cell.DataCell.Type) is StringCellValueWriterBase);
        var result = StringCellValueWriterBase.WriteStartElementWithReference(actualStyleId, state);
        Debug.Assert(result);
    }

    public bool TryWriteValue(in StyledCell cell, ref int valueIndex, CellWriterState state)
    {
        return state.Buffer.WriteLongStringXmlEncoded(cell.DataCell.Value.StringOrPrimitive.StringValue, ref valueIndex);
    }

    public bool TryWriteEndElement(in StyledCell cell, CellWriterState state)
    {
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return writer.TryWriteEndElement(state.Buffer);
    }
}