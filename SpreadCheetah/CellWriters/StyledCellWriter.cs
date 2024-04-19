using SpreadCheetah.CellValueWriters;
using SpreadCheetah.Styling.Internal;

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

    protected override bool WriteStartElement(in StyledCell cell)
    {
        return cell.StyleId is null
            ? CellValueWriter.GetWriter(cell.DataCell.Type).WriteStartElement(Buffer)
            : CellValueWriter.GetWriter(cell.DataCell.Type).WriteStartElement(cell.StyleId, Buffer);
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
