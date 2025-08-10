using SpreadCheetah.CellValueWriters;
using SpreadCheetah.CellValueWriters.Characters;
using SpreadCheetah.Styling;
using SpreadCheetah.Styling.Internal;
using System.Diagnostics;

namespace SpreadCheetah.CellWriters;

internal sealed class StyledCellWriter(CellWriterState state, DefaultStyling? defaultStyling)
    : BaseCellWriter<StyledCell>(state, defaultStyling)
{
    protected override bool TryWriteCell(in StyledCell cell)
    {
        return cell.StyleId is null
            ? CellValueWriter.GetWriter(cell.DataCell.Type).TryWriteCell(cell.DataCell, DefaultStyling, Buffer)
            : CellValueWriter.GetWriter(cell.DataCell.Type).TryWriteCell(cell.DataCell, cell.StyleId, Buffer);
    }

    protected override bool TryWriteCell(in StyledCell cell, StyleId styleId)
    {
        var actualStyleId = cell.StyleId ?? styleId;
        return CellValueWriter.GetWriter(cell.DataCell.Type).TryWriteCell(cell.DataCell, actualStyleId, Buffer);
    }

    protected override bool WriteStartElement(in StyledCell cell, StyleId? styleId)
    {
        var actualStyleId = cell.StyleId ?? styleId;
        Debug.Assert(CellValueWriter.GetWriter(cell.DataCell.Type) is StringCellValueWriterBase);
        return StringCellValueWriterBase.WriteStartElement(actualStyleId, Buffer);
    }

    protected override bool TryWriteEndElement(in StyledCell cell)
    {
        return CellValueWriter.GetWriter(cell.DataCell.Type).TryWriteEndElement(Buffer);
    }

    protected override bool FinishWritingCellValue(in StyledCell cell, ref int cellValueIndex)
    {
        return Buffer.WriteLongString(cell.DataCell.Value.StringOrPrimitive.StringValue, ref cellValueIndex);
    }
}
