using SpreadCheetah.Styling.Internal;

namespace SpreadCheetah.CellWriters;

internal sealed class StyledCellWithReferenceWriter(CellWriterState state, DefaultStyling? defaultStyling)
    : BaseCellWriter<StyledCell>(state, defaultStyling)
{
    protected override bool TryWriteCell(in StyledCell cell)
    {
        return cell.StyleId is null
            ? cell.DataCell.Writer.TryWriteCellWithReference(cell.DataCell, DefaultStyling, State)
            : cell.DataCell.Writer.TryWriteCellWithReference(cell.DataCell, cell.StyleId, State);
    }

    protected override bool WriteStartElement(in StyledCell cell)
    {
        return cell.StyleId is null
            ? cell.DataCell.Writer.WriteStartElementWithReference(State)
            : cell.DataCell.Writer.WriteStartElementWithReference(cell.StyleId, State);
    }

    protected override bool TryWriteEndElement(in StyledCell cell)
    {
        return cell.DataCell.Writer.TryWriteEndElement(Buffer);
    }

    protected override bool FinishWritingCellValue(in StyledCell cell, ref int cellValueIndex)
    {
        return Buffer.WriteLongString(cell.DataCell.StringValue, ref cellValueIndex);
    }
}
