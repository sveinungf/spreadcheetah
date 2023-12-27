using SpreadCheetah.CellValueWriters;
using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellWriters;

internal sealed class StyledCellWithReferenceWriter(CellWriterState state, DefaultStyling? defaultStyling)
    : BaseCellWriter<StyledCell>(state, defaultStyling)
{
    protected override bool TryWriteCell(in StyledCell cell)
    {
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return cell.StyleId is null
            ? writer.TryWriteCellWithReference(cell.DataCell, DefaultStyling, State)
            : writer.TryWriteCellWithReference(cell.DataCell, cell.StyleId, State);
    }

    protected override bool WriteStartElement(in StyledCell cell)
    {
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return cell.StyleId is null
            ? writer.WriteStartElementWithReference(State)
            : writer.WriteStartElementWithReference(cell.StyleId, State);
    }

    protected override bool TryWriteEndElement(in StyledCell cell)
    {
        var writer = CellValueWriter.GetWriter(cell.DataCell.Type);
        return writer.TryWriteEndElement(Buffer);
    }

    protected override bool FinishWritingCellValue(in StyledCell cell, ref int cellValueIndex)
    {
        return Buffer.WriteLongString(cell.DataCell.StringValue, ref cellValueIndex);
    }
}
